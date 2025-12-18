import requests
import json

def test_with_new_session():
    """使用新的会话测试以确保获取到计算结果"""
    
    print("使用AI创建一个包含公式的表格...")
    
    # 创建测试数据
    with open('simple_test_data.csv', 'w') as f:
        f.write("Name,Value1,Value2\n")
        f.write("Item1,10,20\n")
        f.write("Item2,30,40\n")
        f.write("Item3,50,60\n")
    
    # 先获取原始数据
    print("获取原始数据...")
    with open('simple_test_data.csv', 'rb') as f:
        files = {'file': f}
        response = requests.post('http://localhost:8082/api/excel/get-data', files=files)
        original_data = response.json()
    
    print("原始数据:")
    for row in original_data['data']:
        print(row)
    
    # 使用AI添加公式
    print("\n向表格添加公式...")
    with open('simple_test_data.csv', 'rb') as f:
        files = {'file': f}
        data = {'command': 'Add a new column called \'Total\' that sums Value1 and Value2 for each row'}
        response = requests.post('http://localhost:8082/api/ai/excel-with-ai', files=files, data=data)
        result = response.json()
    
    print("AI响应:", result['aiResponse'])
    print("命令执行结果:")
    for cmd_result in result['commandResults']:
        print(f"  - {cmd_result['message']}")
    
    # 检查AI生成的数据预览
    print("\nAI处理后的数据预览:")
    print(result['excelDataPreview'])

if __name__ == "__main__":
    test_with_new_session()