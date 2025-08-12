# -*- coding: utf-8 -*-
import pandas as pd
import duckdb
import json
import re
from typing import Optional, Dict, List, Any, Tuple
from .tool_registry import ToolHandler

# --------------------------------------------------------------------------- #
#                      模块一: 底层清洗和加载函数                  #
# --------------------------------------------------------------------------- #

def clean_column_names_with_replacement(df: pd.DataFrame) -> pd.DataFrame:
    new_columns = []
    _seen = set()
    for col in df.columns:
        col_cleaned = str(col).strip()
        col_cleaned = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]+', '_', col_cleaned)
        col_cleaned = col_cleaned.strip('_')
        if not col_cleaned: col_cleaned = "unnamed_column"
        original_col, counter = col_cleaned, 1
        while col_cleaned in _seen:
            col_cleaned = f"{original_col}_{counter}"
            counter += 1
        _seen.add(col_cleaned)
        new_columns.append(col_cleaned)
    df.columns = new_columns
    return df

# --------------------------------------------------------------------------- #
#               模块二: 数据查询与执行工具 (Tool, 职责单一)                   #
# --------------------------------------------------------------------------- #

class DataAnalysisToolMultiTable:
    """一个纯粹的执行工具，接收数据和SQL，返回结果。"""
    def execute_sql(self, dfs: Dict[str, pd.DataFrame], sql_query: str) -> str:
        if not isinstance(dfs, dict) or not all(isinstance(df, pd.DataFrame) for df in dfs.values()):
            return self._format_error("输入参数 'dfs' 必须是一个键为表名，值为Pandas DataFrame的字典。")
        con: Optional[duckdb.DuckDBPyConnection] = None
        try:
            con = duckdb.connect(database=':memory:', read_only=False)
            for table_name, df in dfs.items():
                con.register(table_name, df)
            
            # 分割SQL语句（通过分号分隔）
            sql_statements = [stmt.strip() for stmt in sql_query.split(';') if stmt.strip()]
            
            if len(sql_statements) == 1:
                # 单语句查询，返回原格式
                result_df = con.execute(sql_statements[0]).fetchdf()
                return self._prepare_dataframe_for_json(result_df).to_json(orient='records', date_format='iso', default_handler=str)
            else:
                # 多语句查询，返回结果列表
                results = []
                for i, sql_stmt in enumerate(sql_statements):
                    try:
                        result_df = con.execute(sql_stmt).fetchdf()
                        results.append({
                            "query_index": i + 1,
                            "sql": sql_stmt[:100] + "..." if len(sql_stmt) > 100 else sql_stmt,
                            "result": json.loads(self._prepare_dataframe_for_json(result_df).to_json(orient='records', date_format='iso', default_handler=str)),
                            "row_count": len(result_df),
                            "column_count": len(result_df.columns)
                        })
                    except Exception as e:
                        results.append({
                            "query_index": i + 1,
                            "sql": sql_stmt[:100] + "..." if len(sql_stmt) > 100 else sql_stmt,
                            "error": str(e)
                        })
                return json.dumps({"multiple_queries": True, "results": results}, ensure_ascii=False)
                
        except duckdb.Error as e:
            return self._format_error(f"SQL执行失败: {e}")
        except Exception as e:
            return self._format_error(f"数据处理失败: {e}")
        finally:
            if con: con.close()

    def _prepare_dataframe_for_json(self, df: pd.DataFrame) -> pd.DataFrame:
        df_copy = df.copy()
        for col in df_copy.columns:
            if pd.api.types.is_datetime64_any_dtype(df_copy[col]):
                df_copy[col] = df_copy[col].dt.strftime('%Y-%m-%d %H:%M:%S')
        return df_copy
    
    def _format_error(self, message: str) -> str:
        return json.dumps({"error": message}, ensure_ascii=False)


# --------------------------------------------------------------------------- #
#              模块三: 分析协调器 (Orchestrator, 核心流程管理者)               #
# --------------------------------------------------------------------------- #

class ExcelAnalysisOrchestrator:
    """
    管理从Excel加载到LLM交互的全流程。
    """
    def __init__(self, file_path: str):
        """
        初始化时，自动完成“数据发现与描述”阶段。
        """
        print("--- 协调器初始化：开始数据发现与描述 ---")
        self.file_path = file_path
        self.data_tables: Dict[str, pd.DataFrame] = self._load_and_clean_all_sheets()
        self.tool = DataAnalysisToolMultiTable()
        print("--- 协调器准备就绪 ---\n")

    def _load_and_clean_all_sheets(self) -> Dict[str, pd.DataFrame]:
        """
        加载Excel中的所有Sheet，并对每一个进行清洗。
        """
        try:
            all_sheets_dict = pd.read_excel(self.file_path, sheet_name=None, engine='openpyxl')
        except Exception as e:
            print(f"❌ [ERROR] 无法读取Excel文件: {e}")
            return {}
            
        cleaned_tables = {}
        for sheet_name, df in all_sheets_dict.items():
            print(f"  -> 正在处理Sheet: '{sheet_name}'")
            # 清洗列名
            df_cleaned = clean_column_names_with_replacement(df.copy())
            df_cleaned.dropna(how='all', axis=0, inplace=True)
            df_cleaned.dropna(how='all', axis=1, inplace=True)
            df_cleaned.reset_index(drop=True, inplace=True)
            
            # 使用一个更安全的、适合做表名的key
            table_key = re.sub(r'\W+', '', sheet_name).lower()
            if not table_key: table_key = f"sheet_{len(cleaned_tables)}"
            cleaned_tables[table_key] = df_cleaned
        return cleaned_tables

    def get_llm_context(self) -> str:
        """
        生成一个结构化的字符串，用于给LLM提供上下文。
        """
        if not self.data_tables:
            return "错误：未能从Excel文件中加载任何数据。"
        
        context_parts = ["用户上传的Excel文件内容如下："]
        for table_name, df in self.data_tables.items():
            context_parts.append(f"\n--- 表名: `{table_name}` (来自Sheet: '{df.attrs.get('original_sheet_name', table_name)}') ---")
            
            # 列信息
            columns_info = []
            for col in df.columns:
                dtype = str(df[col].dtype)
                columns_info.append(f'"{col}" ({dtype})')
            context_parts.append("列名和类型: " + ", ".join(columns_info))

            # 数据预览
            context_parts.append("数据预览 (前5行):")
            context_parts.append(df.head(5).to_string(index=False))

        return "\n".join(context_parts)

    def run_analysis(self, sql_query: str) -> str:
        """
        执行"数据查询与执行"阶段。
        """
        print("--- 协调器开始执行分析任务 ---")
        return self.tool.execute_sql(self.data_tables, sql_query)


# --------------------------------------------------------------------------- #
#                    模块四: 工具注册系统集成                                   #
# --------------------------------------------------------------------------- #

class SqlExecutionTool(ToolHandler):
    """SQL执行工具"""
    
    @property
    def name(self) -> str:
        return "execute_sql"
    
    def execute(self, args: Dict[str, Any], context: Dict[str, Any]) -> Tuple[bool, str]:
        """
        执行SQL查询
        
        参数:
            args: 包含sql_query的参数字典
            context: 包含excel_orchestrator的上下文
            
        返回:
            (成功标志, 执行结果)
        """
        excel_orchestrator = context.get("excel_orchestrator")
        if not excel_orchestrator:
            return False, json.dumps({
                "error": "请先上传Excel文件"
            }, ensure_ascii=False)
        
        sql_query = args.get("sql_query", "")
        if not sql_query:
            return False, json.dumps({
                "error": "缺少SQL查询语句"
            }, ensure_ascii=False)
        
        try:
            result = excel_orchestrator.run_analysis(sql_query)
            return True, result
        except Exception as e:
            return False, json.dumps({
                "error": f"SQL执行失败：{str(e)}"
            }, ensure_ascii=False)

