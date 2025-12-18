import requests
import json

def test_formula_format():
    """测试修复后的公式格式"""
    
    print("使用AI创建薪资估算列...")
    # 使用AI添加薪资估算列
    with open('simple_test_data.csv', 'rb') as f:
        files = {'file': f}
        data = {'command': 'Create a new column called \'Age Group\' that categorizes people as \'Young\' if under 30, \'Middle-aged\' if 30-33, and \'Senior\' if over 33'}
        response = requests.post('http://localhost:8082/api/ai/excel-with-ai-download', 
                               files=files, data=data,
                               headers={'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
    
    if response.status_code == 200:
        with open('modified_age_group_file.xlsx', 'wb') as f:
            f.write(response.content)
        print("修改后的文件已下载为 'modified_age_group_file.xlsx'")
    else:
        print(f"下载失败，状态码: {response.status_code}")
        print("响应内容:", response.text)
        return
    
    # 上传修改后的文件并获取其数据
    print("\n获取修改后的年龄分组数据...")
    with open('modified_age_group_file.xlsx', 'rb') as f:
        files = {'file': ('modified_age_group_file.xlsx', f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        response = requests.post('http://localhost:8082/api/excel/get-data', files=files)
        modified_data = response.json()
    
    print("修改后的年龄分组数据:")
    for row in modified_data['data']:
        print(row)
    
    print("\n现在测试是否公式能正确生成...")
    print("如果公式以'='开头而不是'=='开头，修复成功！")

if __name__ == "__main__":
    test_formula_format()