import requests
import json
import time

def test_column_creation():
    """测试AI是否真正创建了新列"""
    
    # 首先获取原始数据
    print("获取原始数据...")
    with open('simple_test_data.csv', 'rb') as f:
        files = {'file': f}
        response = requests.post('http://localhost:8082/api/excel/get-data', files=files)
        original_data = response.json()
    
    print("原始数据:")
    for row in original_data['data']:
        print(row)
    
    print("\n使用AI添加新列...")
    # 使用AI添加新列
    with open('simple_test_data.csv', 'rb') as f:
        files = {'file': f}
        data = {'command': 'Add a new column called \'Department\' and fill it with appropriate department names based on the person\'s name'}
        response = requests.post('http://localhost:8082/api/ai/excel-with-ai', files=files, data=data)
        ai_result = response.json()
    
    print("AI响应:", json.dumps(ai_result, indent=2))
    
    # 下载修改后的文件
    print("\n下载修改后的文件...")
    with open('simple_test_data.csv', 'rb') as f:
        files = {'file': f}
        data = {'command': 'Add a new column called \'Department\' and fill it with appropriate department names based on the person\'s name'}
        response = requests.post('http://localhost:8082/api/ai/excel-with-ai-download', 
                               files=files, data=data,
                               headers={'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
    
    if response.status_code == 200:
        with open('modified_file.xlsx', 'wb') as f:
            f.write(response.content)
        print("修改后的文件已下载为 'modified_file.xlsx'")
    else:
        print(f"下载失败，状态码: {response.status_code}")
        print("响应内容:", response.text)
        return
    
    # 上传修改后的文件并获取其数据
    print("\n获取修改后的数据...")
    with open('modified_file.xlsx', 'rb') as f:
        files = {'file': ('modified_file.xlsx', f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        response = requests.post('http://localhost:8082/api/excel/get-data', files=files)
        modified_data = response.json()
    
    print("修改后的数据:")
    for row in modified_data['data']:
        print(row)

if __name__ == "__main__":
    test_column_creation()