# Mac版本运行说明

## 首次运行方法

由于macOS的安全机制，从GitHub下载的应用首次运行时会被Gatekeeper阻止。请按以下步骤操作：

### 方法一：右键打开

1. 下载 `PPTgenerator`文件到本地
2. **右键点击**文件，选择"打开"
3. 在弹出的对话框中点击"打开"按钮
4. 之后就可以正常双击运行了

### 方法二：首次终端运行（推荐），后续可正常双击程序打开

1. 打开终端（Terminal）
2. 进入文件所在目录：`cd 这里需要复制你的文件路径，你可以在程序所在目录的下方路径的最后一个文件夹名字上右击，“复制为文件路径”相关的描述`
3. 允许执行：`chmod +x PPTgenerator`
4. 运行：`./PPTgenerator`

### 方法三：系统设置（如果上述方法不行）

1. 打开"系统偏好设置" → "安全性与隐私" → "通用"
2. 如果看到关于PPTgenerator被阻止的提示，点击"仍要打开"
3. 确认打开应用

## 版本选择

现在提供两个Mac版本，请根据您的Mac芯片选择：

- **PPTgenerator-intel**: 适用于Intel芯片Mac（2020年之前的Mac）
- **PPTgenerator-arm64**: 适用于Apple Silicon芯片Mac（M1/M2/M3等，2020年后的Mac）

## 常见问题

**Q: 提示"未知开发者"怎么办？**
A: 使用上述"方法一"的右键打开方式即可绕过此限制。

**Q: 提示"bad CPU type in executable"怎么办？**
A: 说明您下载了错误的版本，请重新下载对应您Mac芯片的版本。

**Q: 提示文件已损坏？**
A: 在终端中运行：`xattr -d com.apple.quarantine /path/to/PPTgenerator-intel`（或PPTgenerator-arm64）

## 注意事项

- 首次运行需要右键打开，之后可以正常双击运行
- 如果移动文件位置，可能需要重新执行首次运行步骤
