# AI 论文阅读生成 PPT — Version 1

一个本地 Web 工具，把论文内容转成可编辑的 PowerPoint 文件。支持 DeepSeek / OpenAI 兼容接口，也可不填接口使用本地摘要逻辑。

## 功能

- 粘贴论文文本，或上传 TXT/MD/PDF
- 支持 DeepSeek / OpenAI 兼容接口地址、模型名和 API Key
- **不限页数**：选择"不限"让 AI 根据论文内容自动决定页数，以讲清楚为准
- **原始数据**：自动提取论文中的实验数据（准确率、精度、召回等指标）嵌入幻灯片
- **智能图示**：每页自动规划流程图/柱状图/矩阵图/时间线图
- 页面内显示生成进度条
- 生成 `.pptx` 并提供下载链接
- 未填写接口或接口失败时，自动使用本地摘要逻辑生成基础 PPT

## 接入 DeepSeek

页面已预填 DeepSeek 官方地址：

```text
https://api.deepseek.com
```

你只需要填写 API Key 并选择模型：

- 获取 Key：`https://platform.deepseek.com/api_keys`
- `deepseek-v4-flash`：推荐，速度快
- `deepseek-v4-pro`：质量更高，适合长论文
- `deepseek-chat` / `deepseek-reasoner`：旧兼容名（2026-07-24 后废弃）

## 启动

```powershell
python app.py
```

打开 `http://127.0.0.1:8765`

## API

| 接口 | 说明 |
|------|------|
| `GET /` | 页面入口 |
| `POST /api/generate` | 提交论文生成 PPT（multipart/form-data） |
| `GET /api/jobs/{job_id}` | 查询进度 |
| `GET /api/download/{job_id}` | 下载生成的 PPTX |

### `/api/generate` 字段

| 字段 | 类型 | 说明 |
|------|------|------|
| `paper_text` | string | 论文文本 |
| `paper_file` | file | TXT/MD/PDF（可选） |
| `api_base` | string | 接口地址 |
| `api_key` | string | API Key |
| `model` | string | 模型名 |
| `slides` | int | 页数（0 = 不限，让 AI 自动决定） |
| `language` | string | 中文 / English |

生成结果保存在 `outputs/` 目录。

## Version 1 更新

- 移除页数硬限制（原最大 12 页）
- 新增"不限"模式，AI 按内容自动决定页数
- AI 提示词优化，要求包含论文原始实验数据
- 本地模式新增数据点提取（`extract_data_points`），自动识别含数值指标的行
- 本地模式 slide_count=0 时按论文内容丰富度自动生成
