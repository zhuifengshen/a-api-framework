import requests

'''
演进0：纯脚本接口自动化测试
'''

# 一、首页接口
# 建立url_index的变量，存储战场的首页
url_index = 'http://127.0.0.1:12306/'
# 调用requests类的get方法，也就是HTTP的GET请求方式，访问了url_index存储的首页URL，返回结果存到了response_index中
response_index = requests.get(url_index)
# 存储返回的response_index对象的text属性存储了访问主页的response信息，通过下面打印出来
print('Response内容：'+response_index.text)

# 二、登录接口
# 建立url_login的变量，存储战场系统的登录URL
url_login = 'http://127.0.0.1:12306/login'
# username变量存储用户名参数
username = 'devin'
# password变量存储密码参数
password = 'devin'
# 拼凑body的参数
payload = 'username=' + username + '&password=' + password
# 调用requests类的post方法，也就是HTTP的POST请求方式，
# 访问了url_login，其中通过将payload赋值给data完成body传参
response_login = requests.post(url_login, data=payload)
# 存储返回的response_index对象的text属性存储了访问主页的response信息，通过下面打印出来
print('Response内容：' + response_login.text)
