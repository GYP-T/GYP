import xlwings as xw
import time
import random
import re

def insert(file,post,name,today,work):
    # 文件读取
    app = xw.App()
    wb = xw.books.open(file)
    sht = wb.sheets['Sheet1']
    # 科室插入
    sht.range("C2").value = post
    # 姓名插入
    sht.range("G2").value = name
    # 时间插入
    sht.range("E2").value = today
    for i in range(4,12):
        # 工作内容插入
        k = random.randint(0,len(work)-1)
        sht.range("D%d" %i).value = work[k]

    save_file = re.sub('(?<=/)\w+.xlsx','%s%s工作日报.xlsx' %(today,name),file)
    # 文件另存为位置
    wb.save(save_file)
    # 程序退出
    wb.close()
    app.kill()
    print('我运行完了')



if __name__ == '__main__':


    # 文件位置
    file = 'H:/制造科工作日报/办公室工作日报管理制度工作日报表.xlsx'
    # 科室
    post = "制造科"
    # 姓名
    name = "马莹"
    # 当天时间
    today = time.strftime('%Y-%m-%d')
    # 工作内容 可以实际内容填加
    work = ["发带装切受刀","查看设备室相关邮件并做相应回复",
            "购买零部件","查看零部件订购进展","整理月结资金报账",
            "向HDK请求故障品费用","向HDK请求消耗品订购事项",
	   "月底盘点","录入盘点数据","核对盘点数据","处理车间紧急发生事项"
	    ]

    insert(file,post,name,today,work)