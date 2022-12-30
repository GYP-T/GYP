import xlwings as xw
import os
import time



def catalogue(today,name1):
    global app, wb, sht
    try:
        # 确认日期文件夹是否存在，防止数据擦除
        if  os.path.exists("H:/制造科工作日报/%s" %today):
            # 创建空白EXCEL文件
            app = xw.App()
            wb = xw.Book()
            # 创建日期文件夹
            for k,v in name1.items():
                # 确认部门文件夹是否创建
                if  os.path.exists("H:/制造科工作日报/%s/%s" % (today,k)):

                    # 部门内人员姓名迭代
                    for i in v:
                        # 创建人员姓名文件夹
                        file = os.listdir("H:/制造科工作日报/%s/%s/%s" % (today,k,i))

                        if  len(file) != 0 :
                            # 消除文本中的''
                            file_name = file[0].strip('')
                            # 拼接连接地址
                            file_path = "H:/制造科工作日报/%s/%s/%s/%s" % (today,k,i,file_name)
                            # 创建对应姓名工作表
                            sht = xw.sheets.add(i)
                            # 调用数据粘贴函数
                            insert(file_path)
                            # 调用格式处理函数
                            style()
            # 文件保存
            wb.save("H:/制造科工作日报/%s/%s制造科日报.xlsx" %(today,today))
            # 关闭文件
            wb.close()
            app.kill()
    except Exception as e:
        print("请检查日期")

    print("创建完成")



def insert(file):

    # 文件读取
    app1 = xw.App()
    wb1 = xw.books.open(file)
    sht1 = wb1.sheets['Sheet1']
    sht1.range("B1:G17").api.Copy()
    sht.range("B1").api.Select()
    sht.api.Paste()
    # 剪切板上有大量信息，是否是要保存  = False
    wb1.app.api.CutCopyMode = False
    wb1.close()
    app1.kill()




    # # 科室插入
    # sht.range("C2").value = post
    # # 姓名插入
    # sht.range("G2").value = name
    # # 时间插入
    # sht.range("E2").value = today
    # for i in range(4,12):
    #     # 工作内容插入
    #     k = random.randint(0,len(work)-1)
    #     sht.range("D%d" %i).value = work[k]
    #
    # save_file = re.sub('(?<=/)\w+.xlsx','%s%s工作日报.xlsx' %(today,name),file)
    # # 文件另存为位置
    # wb.save(save_file)
    # # 程序退出
    # wb.close()
    # app.kill()
    # print('工作随机生产完成')


def style():
    # 设置行高
    height = [41,45,45,49,49,49,49,49,49,49,49,44,44,44,39,100,28]
    k = 1
    for i in height:
        sht.range(k,1).row_height = i
        k += 1

    # 设置行宽
    width = [8.38, 16.13, 16.13, 13.88, 13.88, 13.88, 13.88]
    n = 1
    for v in width:
        sht.range(4,n).column_width = v
        n += 1




if __name__ == '__main__':


    # 当天时间
    today = time.strftime('%Y-%m-%d')




    name1 = {"办公室":("金立","王爽"),"设备室":("顾圆鹏","马莹"),"一车间":("丁明亮","张小伟","韩文彬","董克旭","邵杰",
                 "王明达","袁学勇","胡贺鑫","董春德","白洪学","刘宇","李承楠","许晓文","张会生","张凤阳"),
                 "二车间":("刘艳","陈泉宇","彭朝旺","杨志鹏","夏凯旋","时佳鸣")
                 }

    print("欢迎使用文件大聪明整理工具")
    print("请输入要整理的日期（q为本日）")
    print("文件格式为2022-12-01")
    day = str(input("请输入:"))
    if len(today) - len(day) <= 3:
        today = day

    catalogue(today, name1)
    input("按任意键退出程序...........")