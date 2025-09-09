import requests
import time
import json
import csv
from tqdm import tqdm  # 用于显示进度条，可选


def get_stock_info(stock_code):
    """
    获取股票信息

    参数:
        stock_code (str): 股票代码，如"600000"

    返回:
        dict: 包含所有需要字段的字典
    """
    base_url = "https://query.sse.com.cn/commonQuery.do"
    headers = {
        "Host": "query.sse.com.cn",
        "Referer": "https://www.sse.com.cn/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/91.0.4472.124 Safari/537.36"
    }

    # 获取股本数据
    def get_volume_data():
        params = {
            "jsonCallBack": f"jsonpCallback",
            "isPagination": "false",
            "sqlId": "COMMON_SSE_CP_GPJCTPZ_GPLB_GPGK_GBJG_C",
            "COMPANY_CODE": stock_code,
            "_": str(int(time.time() * 1000))
        }
        response = requests.get(base_url, headers=headers, params=params)
        json_str = response.text.replace(params["jsonCallBack"], "").strip("();")
        return json.loads(json_str)

    # 获取公司基本信息
    def get_company_info():
        params = {
            "jsonCallBack": f"jsonpCallback{int(time.time() * 1000)}",
            "isPagination": "false",
            "sqlId": "COMMON_SSE_CP_GPJCTPZ_GPLB_GPGK_GSGK_C",
            "COMPANY_CODE": stock_code,
            "_": str(int(time.time() * 1000))
        }
        response = requests.get(base_url, headers=headers, params=params)
        json_str = response.text.replace(params["jsonCallBack"], "").strip("();")
        return json.loads(json_str)

    try:
        # 获取两部分数据
        volume_data = get_volume_data()
        company_data = get_company_info()
        time.sleep(0.2)  # 避免请求过于频繁

        result = {}

        # 处理股本数据
        if volume_data.get("result") and len(volume_data["result"]) > 0:
            vol_result = volume_data["result"][0]
            total_domestic = vol_result.get("TOTAL_DOMESTIC_VOL", "N/A")
            total_unlimit = vol_result.get("TOTAL_UNLIMIT_VOL", "N/A")

            # 将数值乘以10000并转换为整数（去掉小数点）
            try:
                if total_domestic != "N/A":
                    total_domestic = str(int(float(total_domestic) * 10000))
                if total_unlimit != "N/A":
                    total_unlimit = str(int(float(total_unlimit) * 10000))
            except (ValueError, TypeError):
                pass

            result.update({
                "TOTAL_DOMESTIC_VOL": total_domestic,
                "TOTAL_UNLIMIT_VOL": total_unlimit,
                "TRADE_DATE": vol_result.get("TRADE_DATE", "N/A")
            })

        # 处理公司信息
        if company_data.get("result") and len(company_data["result"]) > 0:
            comp_result = company_data["result"][0]
            result.update({
                "FULL_NAME": comp_result.get("FULL_NAME", "N/A"),
            })

        return result

    except Exception as e:
        print(f"获取股票 {stock_code} 数据时出错: {str(e)}")
        return {"error": str(e)}


def process_csv(input_file):
    """
    处理CSV文件，更新股本数据

    参数:
        input_file (str): 输入CSV文件路径
    """
    # 读取CSV文件
    with open(input_file, 'r', encoding='gbk') as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        fieldnames = reader.fieldnames

    # 处理每一行数据
    for row in tqdm(rows, desc="正在处理股票数据"):
        stock_code = row['A股代码']
        stock_info = get_stock_info(stock_code)

        if "error" not in stock_info:
            # 更新总股本和流通股数据
            row['总股本'] = stock_info.get("TOTAL_DOMESTIC_VOL", row['总股本'])
            row['流通股'] = stock_info.get("TOTAL_UNLIMIT_VOL", row['流通股'])

    # 写回CSV文件
    with open(input_file, 'w', encoding='gbk', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


if __name__ == "__main__":
    csv_file = input("请输入CSV文件路径: ")
    process_csv(csv_file)
    print("数据处理完成并已更新到原文件")
