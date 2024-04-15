import openpyxl
import requests
from urllib.parse import urlparse
import os

def download_and_rename_files_from_excel(excel_file):
    # 设置一个常见的用户代理
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }
    
    # 打开 Excel 文件
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active
    
    # 创建两个文件夹
    os.makedirs('文件夹1', exist_ok=True)
    os.makedirs('文件夹2', exist_ok=True)
    
    # 循环读取第一列的名字和第二三列的超链接
    for row in sheet.iter_rows(min_row=2, max_col=3, values_only=True): # 从第二行开始读取
        name = row[0] # 获取名字
        url1 = row[1] # 获取超链接1
        url2 = row[2] # 获取超链接2
        
        # 获取文件名1
        parsed_url1 = urlparse(url1)
        file_name1 = os.path.basename(parsed_url1.path)
        
        # 获取文件名2
        parsed_url2 = urlparse(url2)
        file_name2 = os.path.basename(parsed_url2.path)
        
        # 下载文件1
        response1 = requests.get(url1, headers=headers)
        
        # 检查响应状态1
        if response1.status_code == 200:
            # 写入文件1
            with open(os.path.join('文件夹1', file_name1), 'wb') as file:
                file.write(response1.content)
            
            # 构建新文件名1
            _, file_extension1 = os.path.splitext(file_name1)
            new_file_name1 = f"Team{name}-1stPaperReview{file_extension1}"
            
            # 重命名文件1
            os.rename(os.path.join('文件夹1', file_name1), os.path.join('文件夹1', new_file_name1))
            
            print(f"文件夹1 {file_name1} 下载并重命名为 {new_file_name1} 完成.")
        else:
            print(f"文件夹1 {file_name1} 下载失败. HTTP状态码：{response1.status_code}")
        
        # 下载文件2
        response2 = requests.get(url2, headers=headers)
        
        # 检查响应状态2
        if response2.status_code == 200:
            # 写入文件2
            with open(os.path.join('文件夹2', file_name2), 'wb') as file:
                file.write(response2.content)
            
            # 构建新文件名2
            _, file_extension2 = os.path.splitext(file_name2)
            new_file_name2 = f"Team{name}-1stPaperReview{file_extension2}"
            
            # 重命名文件2
            os.rename(os.path.join('文件夹2', file_name2), os.path.join('文件夹2', new_file_name2))
            
            print(f"文件夹2 {file_name2} 下载并重命名为 {new_file_name2} 完成.")
        else:
            print(f"文件夹2 {file_name2} 下载失败. HTTP状态码：{response2.status_code}")

# 调用函数并传入 Excel 文件路径
excel_file = '1.xlsx'  # 修改为你的 Excel 文件路径
download_and_rename_files_from_excel(excel_file)
