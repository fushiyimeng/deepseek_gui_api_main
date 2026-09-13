# DeepSeek GUI 聊天窗口

一个基于 PyQt5 的图形界面聊天工具，通过 OpenAI 兼容接口调用 DeepSeek API。

## 功能特性

- 图形界面：会话列表、气泡对话、左右分栏
- 配置外置：API Key、模型、余额等参数保存在本地 `data/config.json`，**不写入代码**
- 按量计费：根据本地时间自动判断高峰/空闲时段，按官方价格计算每次调用的费用
- 用量统计：每次调用记录到 `data/usage.xlsx`（输入命中 / 未命中 / 输出 / 费用）
- 会话持久化：每个会话单独保存为 `data/sessions/*.json`，关闭后重启可继续
- 本地预估（可选）：接入 DeepSeek 官方 tokenizer，输入时实时预估 token 数
- 分卷日志：聊天内容按大小自动分卷写入 `data/chat/*.txt`

## 环境要求

- Python 3.8 及以上

### 依赖

```
PyQt5>=5.15
openai>=1.0
openpyxl>=3.0

# 以下为可选项，用于本地 token 预估
transformers>=4.30
```

安装：

```bash
pip install -r requirements.txt
```

## 使用

1. 启动程序：

   ```bash
   python deespeek.py
   ```

2. 首次使用点右上角「设置」，填写：

   | 字段 | 说明 |
   |---|---|
   | API 名称 | 随便起，仅用于顶栏显示 |
   | Base URL | `https://api.deepseek.com`（不要带 `/v1`） |
   | API Key | 到 DeepSeek 控制台申请，`sk-` 开头 |
   | 模型名 | `deepseek-flash` 或 `deepseek-v4-pro` |
   | 余额(元) | 账上现有金额，程序会按次扣减 |
   | 已用 tokens | 累计值，可手动校准或清零 |

3. 填完点「确定」，就可以在输入框里打字聊天了。Enter 发送，Shift+Enter 换行。

## 计费说明

单价参考 DeepSeek 官网价格，按「元 / 百万 token」计算，分三类：

- 缓存命中输入
- 缓存未命中输入
- 输出

时段划分（依据本地时间）：

- 周一 ~ 周五：`09:00~12:00`、`14:00~18:00` 为高峰，其余为空闲
- 周六、周日：全天为空闲

价格表在代码顶部的 `MODEL_PRICING` 中，如需修改请自行编辑。

## 数据目录

程序首次运行会在当前目录创建 `data/`：

```
data/
├── config.json            配置文件（API、模型、余额、已用 token）
├── usage.xlsx             每次调用的用量与费用明细
├── sessions/              会话内容（每会话一个 JSON）
├── chat/                  聊天文本（按大小自动分卷）
└── logs/app.log           运行日志
```

## 可选：启用本地 token 预估

本地 token 预估功能需要 DeepSeek 官方 tokenizer 文件。启用方法：

1. 从 DeepSeek 官方下载 `deepseek_tokenizer.zip`
2. 解压到程序同级目录，目录名保持为 `deepseek_tokenizer/`
3. 安装依赖：`pip install transformers`

未放置该目录时程序仍可正常运行，仅「预计输入 tokens」一栏不显示数字。

## 免责声明

本项目仅供学习和个人使用。API 调用产生的费用由使用者自行承担，作者不对任何费用、数据丢失或服务中断负责。
