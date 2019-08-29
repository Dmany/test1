from openpyxl import load_workbook
import pandas as pd
import xlsxwriter
import xlrd
import requests
import json

if __name__ == '__main__':

    aa = {}
    aa_id = 0
    for i in range(39):
        workbook = pd.read_excel('test.xlsx')    #打开user.xlsx工作簿
        phone = workbook['phone']   #获取phoen列
        phone1 = phone[i]   #每次循环取1个值

        r = requests.post(  #接口获取ELE余额
            url='http://u.maopai.dev.5chuanmei.cn/api/queryAccountByCode',    #请求ELE接口
            data={"mobile":phone1,"codeName":'ELE'}     #传参
        )
        ser = r.json()  #json转换python对象
        cuwu = ser['errmsg']

        if cuwu =='没有此帐号':
            print(phone1,'进入新注册')
            reginst = requests.post(
                # 接口获取
                url='http://u.maopai.dev.5chuanmei.cn/api/reginst',  # 请求注册接口
                data={"mobile": phone1}  # 传参
            )

            r = requests.post(  # 接口获取ELE余额
                url='http://u.maopai.dev.5chuanmei.cn/api/queryAccountByCode',  # 请求ELE接口
                data={"mobile": phone1, "codeName": 'ELE'}  # 传参
            )
            ser = r.json()  # json转换python对象
            dataList = ser.get('data')  # 截取data字典
            elemoney = dataList['moneyAccount']  # 取moneyAccount值

            if elemoney >= 500.0:   #查询现金券金额
                print("进入查询新ele")
                print(phone1,'已经有现金券',elemoney)
                aa[i] = {phone1: elemoney}
                print(aa,'新')

                # print(aa)
            else:
                print(phone1,"进入新充值")
                c = requests.post(  # 接口充值ELE
                    url='http://u.maopai.dev.5chuanmei.cn/api/transferAit',  # 请求充值接口的URL
                    data={"code": 'ELE', "toMobile": phone1, "fromMobile": '13000000000', "count": '500',
                          "appRef": 'python', "remark": 'tongyi', "orderId": '1'}  # 传参
                )
                csr = c.json()
                print('充值成功！',csr)
            reg = reginst.json()
            zheng = reg['errmsg']

        elif cuwu == '操作成功':    #有账号进入查询
            dataList = ser.get('data')      #截取data字典
            elemoney  = dataList['moneyAccount']    #取moneyAccount值

            if elemoney >= 500.0:   #查询现金券金额
                print(phone1,"进入查询ele")
                print(phone1,'已经有现金券',elemoney)
                aa[i] = {phone1:elemoney}
                print(aa)

            else:
                print(phone1,"进入充值")
                c = requests.post(  # 接口充值ELE
                    url='http://u.maopai.dev.5chuanmei.cn/api/transferAit',  # 请求充值接口的URL
                    data={"code": 'ELE', "toMobile": phone1, "fromMobile": '13000000000', "count": '500',
                          "appRef": 'python', "remark": 'tongyi', "orderId": '1'}  # 传参
                )
                csr = c.json()
                print('充值成功！',csr)