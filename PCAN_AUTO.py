from PCANBasic import *
import time
from sys import *
from threading import *
import xlrd

def choice_msg_type():         #用户选择数据类型
    print("请输入数据类型:")
    print("1.标准帧")
    print("2.扩展帧")
    DATA_TYPE = input()
    if DATA_TYPE == "1":
        DATA_TYPE = 1
    elif DATA_TYPE == "2":
        DATA_TYPE = 2
    else:
        print("输入错误!")
        exit()
    return DATA_TYPE
    
def send(ID,REPEAT,CYCLE,DATA_TYPE,DATA,REPEAT_ALL):
    global pd
    sleep_time=int(CYCLE)/1000       #获取发送周期
    msg = TPCANMsg()
    msg.ID = int(ID,16)
    msg.LEN = 8
    if DATA_TYPE ==1:                           #判断扩展帧与标准帧
        msg.MSGTYPE = PCAN_MODE_STANDARD
    elif DATA_TYPE ==2:
        msg.MSGTYPE = PCAN_MODE_EXTENDED
    list_tmp_list=[]
    remark=[]                                     #备注信息集合
    
    

    # 此处的for循环主要是为了将DATA里的所有数据以list形式转化为tuple格式
    # 后续方便传入msg.DATA
    for data_1_8 in DATA:                         #获取DADA表格每行所有的数据
        list_tmp_2_list=[]
        for data_1_8_every in data_1_8:                            #获取DATA表格每行1~8的数据,并创建集合
            try:          #将每行中的数据和备注信息分离
                list_tmp_2_list.append(int(data_1_8_every, 16))
            except:            #将备注信息插入备注信息的集合
                remark.append(data_1_8_every)
        list_tmp_2_tuple=tuple(list_tmp_2_list)        #将DATA每行的list转化为tuple
        list_tmp_list.append(list_tmp_2_tuple)          #将转化后的tuple插入统一的list
    list_tmp_tuple=tuple(list_tmp_list)                #将所有的data转化为tuple,后续方便用for循环遍历插入msg.DATA
    rep=0                 #正在进行循环的打印变量信息
    
    
    
    
     # 此处数大循环
    for repert_all in range(int(REPEAT_ALL)):
        rep=rep+1
        for i in range(len(DATA)):
            msg.DATA = (list_tmp_tuple[i])
            msg_show = ""
            for x in range(8):
                msg_show = msg_show + " " + ('%#x' % msg.DATA[x])
                
            print(('%#x' % msg.ID) + ':' + msg_show + "      " + CYCLE + "ms","                    Remark:",remark[i])
            print("单帧重复发送次数:",int(REPEAT))
            print()
            print("总计",int(REPEAT_ALL),"次循环",",正在发送第",rep,"次循环")
            print("--------------------------------------------------------------------------------------------")
            for repeat in range(int(REPEAT)):   #此处小循环,每条消息发送几遍
                pd.Write(PCAN_USBBUS1, msg)
                # print("----------------------------------------------------------")
                # print(('%#x'%msg.ID) + ':' + msg_show+"      "+CYCLE+"ms")
                time.sleep(sleep_time)
                
                
                
def run():
    global pd
    msg_data_director = {}    #每个sheet里面包含的所有数据
    msg_ID_director = {}       #每个sheet里面的ID数据
    msg_CYCLE_director = {}     #每个sheet里面包含的周期数据
    msg_data_use_director = {}    #所有8字节数据
    msg_threading_director = {}    #多线程词典
    msg_repeat_director={}      #   小循环
    msg_repeat_director_ALL = {}    #大循环
    threads=[]                     #threading集合,后面用来结束
    file_name = "DATA.xlsx"
    try:
        data = xlrd.open_workbook(filename=file_name)
        #
    except Exception as e:
        print(repr(e))
        print("-----------------------------------------------------")
        print("|缺少DATA文件,请放入DATA.xlsx文件后重试!(按回车退出)|")
        print("-----------------------------------------------------")
        input()
        exit()
    pd = PCANBasic()
    pd.Initialize(PCAN_USBBUS1, PCAN_BAUD_250K)
    DATA_TYPE = choice_msg_type()
    msg_number = len(data.sheet_names())  # 获取DATA表格里面有几个sheet
    # table=data.row_values(1)
    for sheet_index in range(msg_number):
        msg_name = "msg" + str(sheet_index + 1)
        msg_data_director[msg_name] = data.sheets()[sheet_index]
        msg_ID = "msg" + str(sheet_index + 1) + "_ID"
        msg_Cycle = "msg" + str(sheet_index + 1) + "_Cycle"
        msg_data = "msg" + str(sheet_index + 1) + "_Data"
        msg_repeat= "msg"+ str(sheet_index+1)+"_Repeat"
        msg_repeat_all = "msg" + str(sheet_index + 1) + "_Repeat_all"
        for every_data in msg_data_director:
            try:
                msg_repeat_director[msg_repeat] = msg_data_director[every_data].row_values(0)[0].split("=")[1]
                # print(msg_data_director[every_data].row_values(0)[0])
            except Exception as REPEAT:
                print(repr(REPEAT))
                print("数据REPEAT格式错误,请修改后重试!")
                input()
                exit()
            try:
                msg_repeat_director_ALL[msg_repeat_all] = msg_data_director[every_data].row_values(3)[0].split("=")[1]
                # print(msg_data_director[every_data].row_values(0)[0])
            except Exception as REPEAT_ALL:
                print(repr(REPEAT_ALL))
                print("数据REPEAT_ALL格式错误,请修改后重试!")
                input()
                exit()
            try:
                msg_ID_director[msg_ID] = msg_data_director[every_data].row_values(1)[0].split("=")[1]
            except Exception as ID:
                print(repr(ID))
                print("数据ID格式错误,请修改后重试!")
                input()
                exit()
            try:
                msg_CYCLE_director[msg_Cycle] = msg_data_director[every_data].row_values(2)[0].split("=")[1]
            except Exception as CYCLE:
                print(repr(CYCLE))
                print("数据CYCLE格式错误,请修改后重试!")
                input()
                exit()
            for book_long in range(msg_data_director[every_data].nrows):
                if book_long == 0:
                    msg_data_use_director[msg_data] = [msg_data_director[every_data].row_values(book_long)]
                else:
                    msg_data_use_director[msg_data].append(msg_data_director[every_data].row_values(book_long))
            try:
                msg_data_use_director[msg_data] = msg_data_use_director[msg_data][4:]
            except Exception as DATA:
                print(repr(DATA))
                print("数据DATA格式错误,请修改后重试!")
                input()
                exit()
    for do_threading in range(msg_number):
        REPEAT = msg_repeat_director["msg" + str(do_threading + 1) + "_Repeat"]
        ID = msg_ID_director["msg" + str(do_threading + 1) + "_ID"]
        DATA = msg_data_use_director["msg" + str(do_threading + 1) + "_Data"]
        CYCLE = msg_CYCLE_director["msg" + str(do_threading + 1) + "_Cycle"]
        threading_name = "threading" + str(do_threading + 1)
        REPEAT_ALL=msg_repeat_director_ALL["msg" + str(do_threading + 1) + "_Repeat_all"]
        msg_threading_director[threading_name] = Timer(0.1, send, (ID, REPEAT,CYCLE, DATA_TYPE, DATA,REPEAT_ALL))
    for x in msg_threading_director:
        msg_threading_director[x].start()
        threads.append(msg_threading_director[x])
    for end in threads:
        end.join()
    pd.Uninitialize(PCAN_USBBUS1)
    print("主线程结束!按回车键退出!")
    input()
    # print(msg_repeat_director_ALL)
    
    
    
if __name__ == "__main__":
    run()
    # input("运行结束,按回车退出")