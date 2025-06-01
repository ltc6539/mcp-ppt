# PPT maker MCP server

🌐 [英文版README](README.md)

这个 MCP server 支持动态创建、编辑和保存PowerPoint演示文稿。它基于[MCP](https://github.com/modelcontextprotocol/python-sdk)构建，并使用[python-pptx](https://python-pptx.readthedocs.io/en/latest/)库，为大模型提供了一个灵活的工具包来添加幻灯片、图像、表格和其他元素。用户只需与大语言模型聊天，就能轻松地制作、编辑和保存演示文稿，简化了整个工作流程。

## 功能特点

- **创建演示文稿**  
  使用标题初始化新的PowerPoint演示文稿，生成唯一的演示文稿ID。

- **幻灯片操作**  
  - **标题幻灯片：** 添加带有可选副标题的标题幻灯片。
  - **内容幻灯片：** 创建带有标题和项目符号内容的幻灯片。
  - **分节幻灯片：** 插入具有居中大标题和可选背景颜色的分节幻灯片。
  - **图像幻灯片：** 添加包含来自本地文件或URL的图像的幻灯片，带有标题和描述性替代文本。
  - **表格幻灯片：** 插入包含定义好的表头和行数据的表格幻灯片。

- **演示文稿管理**  
  - **保存演示文稿：** 将演示文稿写入指定的文件路径，如有需要可处理临时目录。
  - **下载链接：** 生成包含base64编码演示文稿内容的数据URI，用于直接下载。
  - **演示文稿信息：** 检索有关演示文稿的元数据，如幻灯片数量和可用的幻灯片布局。
  - **演示文稿大纲：** 通过专用资源端点获取演示文稿结构的文本大纲。
  - **删除幻灯片：** 通过1为基础的索引删除幻灯片。
  - **导出为Base64：** 将完整演示文稿导出为base64编码的字符串，以便进一步处理。

## 安装

1. **克隆仓库**  
   ```bash
   git clone https://github.com/ltc6539/mcp-ppt.git
   cd mcp-ppt
   ```

2. **创建虚拟环境（可选但推荐）**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # 在Windows上使用: .venv\Scripts\activate
   ```

3. **然后将MCP添加到项目依赖中**
   ```bash
   uv add "mcp[cli]"
   uv run mcp
   ```

您可以在[Claude桌面版](https://claude.ai/download)中安装此服务器，并通过运行以下命令立即与之交互：
```bash
mcp install server-local.py
```

或者，您可以使用MCP检查器进行测试：
```bash
mcp dev server-local.py
```

如果Claude桌面版出现错误，您可能需要在命令字段中输入uv可执行文件的完整路径。您可以在MacOS/Linux上运行`which uv`或在Windows上运行`where uv`来获取此路径。
在启动过程中，服务器会将Python和python-pptx版本信息记录到stderr。执行过程中的任何错误也会打印到stderr，以便于调试。

## 工具列表

每个MCP工具函数都可以通过MCP服务器直接访问。以下是可用的主要操作：

### 1. 创建演示文稿
- **函数：** `create_presentation(title: str) -> str`  
- **描述：** 初始化新的演示文稿并返回唯一的演示文稿ID。
  
### 2. 添加标题幻灯片
- **函数：** `add_title_slide(prs_id: str, title: str, subtitle: Optional[str] = None) -> str`  
- **描述：** 向指定的演示文稿添加标题幻灯片。

### 3. 添加内容幻灯片
- **函数：** `add_content_slide(prs_id: str, title: str, content: List[str]) -> str`  
- **描述：** 插入带有标题和多个项目符号的内容幻灯片。

### 4. 添加分节幻灯片
- **函数：** `add_section_slide(prs_id: str, section_title: str, background_color: Optional[str] = None) -> str`  
- **描述：** 创建具有可自定义背景颜色和居中大文本的分节幻灯片。

### 5. 添加图像幻灯片
- **函数：** `add_image_slide(prs_id: str, title: str, image_path: str, image_description: str) -> str`  
- **描述：** 添加图像幻灯片。图像可以从本地文件加载或从URL下载。

### 6. 添加表格幻灯片
- **函数：** `add_table_slide(prs_id: str, title: str, headers: List[str], rows: List[List[str]]) -> str`  
- **描述：** 插入包含由列标题和数据行定义的表格的幻灯片。

### 7. 保存演示文稿
- **函数：** `save_presentation(prs_id: str, output_path: str) -> str`  
- **描述：** 将演示文稿保存到指定的输出路径，如有必要管理临时目录。

### 8. 获取演示文稿下载链接
- **函数：** `get_presentation_download_link(prs_id: str) -> str`  
- **描述：** 返回包含演示文稿base64编码数据的数据URI，用于直接浏览器下载。

### 9. 获取演示文稿信息
- **函数：** `get_presentation_info(prs_id: str) -> str`  
- **描述：** 检索演示文稿元数据，如幻灯片数量和可用的幻灯片布局详情。

### 10. 获取演示文稿大纲
- **资源端点：** `presentation://{prs_id}/outline`  
- **描述：** 提供演示文稿结构的文本表示，包括幻灯片标题和内容摘要。

### 11. 删除幻灯片
- **函数：** `remove_slide(prs_id: str, slide_index: int) -> str`  
- **描述：** 通过其1为基础的索引从演示文稿中删除幻灯片。

### 12. 导出为Base64
- **函数：** `export_to_base64(prs_id: str) -> str`  
- **描述：** 将演示文稿导出为base64编码的字符串（显示前100个字符作为样本）。

### 13. SVG生成器提示函数
- **函数：** `svggenerator_prompt(description: str) -> list[base.Message]`
- **描述：** 创建一个提示，指示Claude基于自然语言描述生成SVG图像。该函数返回两条消息的列表：
  1. 一条将Claude角色设定为SVG专家的系统消息
  2. 一条包含特定SVG请求的用户消息

### 14. 生成SVG函数
- **函数：** `generate_svg(prs_id: str, svg_markup: str, title: str = None, width: float = 6.0) -> str`
- **描述：** 将SVG标记添加到PowerPoint演示文稿中：
  - 需要演示文稿ID和SVG标记
  - 可选接受标题和宽度参数（默认6英寸）
  - 将SVG写入临时文件
  - 使用rsvg-convert工具将SVG转换为PNG
  - 在演示文稿中创建新幻灯片
  - 如果提供了标题，则将标题添加到幻灯片中
  - 定位并将PNG图像添加到幻灯片
  - 清理临时文件
  - 返回带有幻灯片位置的确认消息

## 错误处理和调试

- **错误检查：**  
  每个工具都验证输入（例如，验证演示文稿ID或文件存在性）并返回描述性错误消息。
  
- **临时目录：**  
  服务器确保文件保存在可写目录（通常是`/tmp`），如果提供的路径是只读的，则会相应地回退。

- **日志记录：**  
  错误和版本信息输出到stderr，以帮助调试和监控。

## 贡献

欢迎贡献。如果您遇到问题或有改进建议，请开启一个问题或提交拉取请求。