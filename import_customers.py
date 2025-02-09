import os
import csv
import stripe
import pandas as pd
import phonenumbers
from dotenv import load_dotenv
from datetime import datetime
from tqdm import tqdm

# 版本信息
__version__ = "1.0.0"

# 加载环境变量
load_dotenv()

# 配置Stripe API密钥
stripe.api_key = os.getenv('STRIPE_SECRET_KEY')

# 设置日志文件夹和文件名
logs_dir = os.path.join(os.path.dirname(__file__), 'logs')
os.makedirs(logs_dir, exist_ok=True)
log_filename = os.path.join(logs_dir, f'import_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt')

def log_error(message):
    """记录错误信息到日志文件"""
    with open(log_filename, 'a', encoding='utf-8') as log_file:
        log_file.write(f"{datetime.now()}: {message}\n")

def create_stripe_customer(row):
    """创建Stripe客户记录"""
    try:
        # 准备客户数据
        customer_data = {
            'name': row['名称'],
            'description': f"订单号: {row['订单编号']}",
            'metadata': {
                'order_number': row['订单编号'],
                'pickup_time': str(row['Pick-up time (local)']),
                'pickup_location': str(row['Pick-up location']),
                'dropoff_time': str(row['Drop-off time (local)']),
                'dropoff_location': str(row['Drop-off location']),
                'duration_days': str(row['Duration (days)']),
                'additional_services': str(row['Additional services'])
            }
        }

        # 添加电话号码
        if pd.notna(row['电话']):
            # 格式化电话号码为E.164标准格式
            phone = str(row['电话'])
            # 如果电话号码以数字开头（没有+号），添加+号
            if phone and phone[0].isdigit():
                phone = '+' + phone
            # 移除所有非数字和加号字符
            phone = ''.join(char for char in phone if char.isdigit() or char == '+')
            customer_data['phone'] = phone

        # 如果有有效的邮箱地址，添加到客户数据中
        if pd.notna(row['邮箱']) and isinstance(row['邮箱'], str) and '@' in row['邮箱']:
            customer_data['email'] = row['邮箱']

        # 创建新客户记录
        customer = stripe.Customer.create(**customer_data)
        return True, customer.id, row['订单编号']

    except stripe.error.StripeError as e:
        error_message = f"导入客户失败 - 订单编号: {row['订单编号']}, 错误: {str(e)}"
        log_error(error_message)
        return False, error_message, row['订单编号']

def main():
    # 检查Stripe API密钥是否配置
    if not stripe.api_key:
        print("错误: 未找到Stripe API密钥。请在.env文件中设置STRIPE_SECRET_KEY。")
        return

    # 获取Excel文件路径
    excel_file = input("请输入Excel文件路径 (.xlsx 或 .xls): ")
    if not os.path.exists(excel_file):
        print(f"错误: 文件 {excel_file} 不存在")
        return
    
    # 验证文件扩展名
    if not excel_file.lower().endswith(('.xlsx', '.xls')):
        print("错误: 请提供有效的Excel文件 (.xlsx 或 .xls)")
        return

    # 读取Excel文件
    try:
        df = pd.read_excel(excel_file)
        # 定义Excel列名与程序中使用的列名的映射关系
        column_mapping = {
            'Customer name': '名称',
            'Klook booking reference ID': '订单编号',
            'Customer phone number': '电话',
            'Customer email': '邮箱',
            'Pick-up time (local)': 'Pick-up time (local)',
            'Pick-up location': 'Pick-up location',
            'Drop-off time (local)': 'Drop-off time (local)',
            'Drop-off location': 'Drop-off location',
            'Duration (days)': 'Duration (days)',
            'Additional services': 'Additional services'
        }
        
        # 检查必需列是否存在
        missing_columns = [col for col in column_mapping.keys() if col not in df.columns]
        if missing_columns:
            print(f"错误: Excel文件缺少以下必需列: {', '.join(missing_columns)}")
            return

        # 重命名列以匹配程序中使用的列名
        df = df.rename(columns=column_mapping)
        required_columns = ['名称', '订单编号', '电话', '邮箱', 'Pick-up time (local)', 'Pick-up location', 'Drop-off time (local)', 'Drop-off location', 'Duration (days)', 'Additional services']
        
        # 检查必需列是否存在
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"错误: CSV文件缺少以下必需列: {', '.join(missing_columns)}")
            return

        success_count = 0
        error_count = 0
        total_records = len(df)

        print(f"开始导入 {total_records} 条记录...")
        
        # 使用tqdm创建进度条
        with tqdm(total=total_records, desc="导入进度", unit="条") as pbar:
            # 逐条处理客户数据
            for index, row in df.iterrows():
                success, result, order_number = create_stripe_customer(row)
                if success:
                    success_count += 1
                else:
                    error_count += 1
                pbar.update(1)

        # 打印导入结果摘要
        print(f"\n导入完成!")
        print(f"成功导入: {success_count} 条记录")
        print(f"失败记录: {error_count} 条")
        if error_count > 0:
            print(f"详细错误信息请查看日志文件: {log_filename}")

    except Exception as e:
        print(f"处理CSV文件时发生错误: {str(e)}")
        log_error(f"处理CSV文件时发生错误: {str(e)}")

if __name__ == '__main__':
    main()