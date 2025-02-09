# Stripe客户导入工具

这是一个用于将Excel文件中的客户数据批量导入到Stripe平台的Python工具。

## 功能特点

- 支持从Excel文件(.xlsx或.xls)批量导入客户数据到Stripe
- 自动处理电话号码格式化为E.164标准
- 支持导入客户基本信息和元数据
- 详细的错误日志记录
- 导入进度实时显示

## 安装要求

- Python 3.6+
- 依赖包：请查看requirements.txt

## 安装步骤

1. 克隆仓库：
```bash
git clone [仓库URL]
cd stripe-customers-import
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

3. 配置环境变量：
- 复制.env.example文件为.env
- 在.env文件中设置你的Stripe API密钥：
```
STRIPE_SECRET_KEY=your_stripe_secret_key
```

## 使用方法

1. 准备Excel文件，确保包含以下必需列：
- Customer name
- Klook booking reference ID
- Customer phone number
- Customer email
- Pick-up time (local)
- Pick-up location
- Drop-off time (local)
- Drop-off location
- Duration (days)
- Additional services

2. 运行导入脚本：
```bash
python import_customers.py
```

3. 根据提示输入Excel文件路径

4. 等待导入完成，查看导入结果摘要

## 错误处理

- 所有导入错误都会被记录到logs目录下的日志文件中
- 日志文件名格式：import_log_YYYYMMDD_HHMMSS.txt

## 注意事项

- 请确保Excel文件中的数据格式正确
- 电话号码会被自动格式化为E.164标准格式
- 邮箱地址必须包含@符号才会被导入

## 许可证

本项目采用GNU General Public License v3.0许可证。详情请参见LICENSE文件。