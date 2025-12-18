import requests
import json

def test_salary_column():
    """测试AI创建薪资估算列"""
    
    print("使用AI创建薪资估算列...")
    # 使用AI添加薪资估算列
    with open('simple_test_data.csv', 'rb') as f:
        files = {'file': f}
        data = {'command': 'Create a new column called \'Salary Estimate\' based on the person\'s age, with values ranging from 40000 to 80000'}
        response = requests.post('http://localhost:8082/api/ai/excel-with-ai-download', 
                               files=files, data=data,
                               headers={'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
    
    if response.status_code == 200:
        with open('modified_salary_file.xlsx', 'wb') as f:
            f.write(response.content)
        print("修改后的文件已下载为 'modified_salary_file.xlsx'")
    else:
        print(f"下载失败，状态码: {response.status_code}")
        print("响应内容:", response.text)
        return
    
    # 上传修改后的文件并获取其数据
    print("\n获取修改后的薪资数据...")
    with open('modified_salary_file.xlsx', 'rb') as f:
        files = {'file': ('modified_salary_file.xlsx', f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        response = requests.post('http://localhost:8082/api/excel/get-data', files=files)
        modified_data = response.json()
    
    print("修改后的薪资数据:")
    for row in modified_data['data']:
        print(row)

if __name__ == "__main__":
    test_salary_column()