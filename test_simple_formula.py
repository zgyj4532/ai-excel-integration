import requests
import json

def test_simple_formula():
    """测试简单公式的计算结果"""
    
    print("使用AI添加一个简单的数学公式...")
    # 创建一个包含简单数值的CSV
    with open('simple_test_data.csv', 'w') as f:
        f.write("Name,Value1,Value2\n")
        f.write("Item1,10,20\n")
        f.write("Item2,30,40\n")
        f.write("Item3,50,60\n")
    
    # 使用AI添加一个简单的加法公式
    with open('simple_test_data.csv', 'rb') as f:
        files = {'file': f}
        data = {'command': 'Add a new column called \'Total\' that sums Value1 and Value2 for each row'}
        response = requests.post('http://localhost:8082/api/ai/excel-with-ai-download', 
                               files=files, data=data,
                               headers={'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
    
    if response.status_code == 200:
        with open('modified_simple_formula.xlsx', 'wb') as f:
            f.write(response.content)
        print("修改后的文件已下载为 'modified_simple_formula.xlsx'")
    else:
        print(f"下载失败，状态码: {response.status_code}")
        print("响应内容:", response.text)
        return
    
    # 获取修改后的数据
    print("\n获取修改后的数据...")
    with open('modified_simple_formula.xlsx', 'rb') as f:
        files = {'file': ('modified_simple_formula.xlsx', f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        response = requests.post('http://localhost:8082/api/excel/get-data', files=files)
        modified_data = response.json()
    
    print("修改后的数据:")
    for row in modified_data['data']:
        print(row)

if __name__ == "__main__":
    test_simple_formula()