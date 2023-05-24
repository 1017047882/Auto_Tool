import pyautogui
import time
import xlrd
#剪贴复制
import pyperclip
#定义鼠标事件
Dir = r"WindowsRPA\Config"

#pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159

def mouseClick(clickTimes,LorR,img,reTry):
    """
    该函数接受以下参数：点击次数、使用哪个鼠标按钮（左键或右键）、点击前要搜索的图像以及未找到图像时重试的次数。
    
    :param clickTimes: 鼠标应单击的次数。
    :param LorR: 用于指定鼠标单击是左键单击还是右键单击
    :param img: 它是一个变量，表示需要在其上执行鼠标单击的图像。
    :param reTry: 用于确定函数在放弃之前应尝试单击鼠标的次数。它可用于处理由于技术问题或意外错误导致鼠标单击失败的情况。
    """
    if reTry > 1:
            i=1
            while i < reTry+1:
                location=pyautogui.locateCenterOnScreen(f'{Dir}\{img}',confidence=0.9)
                if location is not None:
                    pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.3,duration=0.2,button = LorR)
                    print("重复")
                    i+=1
                time.sleep(0.1)
    else:
        while True:
            location = pyautogui.locateCenterOnScreen(f'{Dir}\{img}',confidence=0.9)
            if location is not None:
                #interval：每次点击间隔时间 duration：执行此操作时间 移动速度 省略后瞬间移动完成
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.3,duration=0.2,button=LorR)
                if reTry == 1:
                    break
            print("未找到匹配图片,0.1s后重试")
            time.sleep(0.1)
# 数据检测
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
# ctype     空：0
#           字符串：1
#           数字：2
#           日期：3
#           布尔：4
#           error：5
def dataCheck(sheet):
    checkCmd = True
    #行数检查
    if sheet.nrows<2:
        print ("无相关配置")
        checkCmd = False
    #逐行进行数据检查
    i= 2
    while i < sheet.nrows:
        #第一列 操作类型检查
        cmdType = sheet.row(i)[0]
        if cmdType.ctype != 2 or cmdType.value not in [1,2,3,4,5,6]:
            print (f"第{i}行数据异常，请确认数据！！！")
            checkCmd = False
        #第二列 内容检查
        cmdValue = sheet.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value in [1,2,3] and cmdValue.ctype != 1:
            print (f"第{i}行，第一列数据异常，图片路径不是字符串类型数据！！！")
            checkCmd = False
        #输入类型，内容不能为空
        if cmdType.value == 4 and cmdValue.ctype == 0:
            print (f"第{i}行，第2列数据异常,输入命令参数不能为空！！！")
            checkCmd = False
        #等待类型，内容必须为数字
        if cmdType.value == 5 and cmdValue.ctype != 2:
            print (f"第{i}行，第2列数据异常,等待命令参数必须是数字！！！")
            checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6 and cmdValue.ctype != 2:
            print('第',i,"行,第2列数据异常，滚轮命令参数必须是数字！！！")
            checkCmd = False
        i += 1
    return checkCmd
#任务
def MainWork(img):
    i=1
    while i<st.nrows:
        #获取本行操作类型
        cmdType = st.row(i)[0]
        if cmdType.value == 1:
            #1 单击左键
            #取图片名称
            img = st.row(i)[1].value
            reTry = 1
            if st.row(i)[2].ctype == 2 and st.row(i)[2].value != 0:
                reTry = st.row(i)[2].value

            mouseClick(1,"left",img,reTry)
            print("单击左键",img)
        elif cmdType.value == 2:
            #2 双击左键
            img = st.row(i)[1].value
            reTry = 1
            if st.row(i)[2].ctype == 2 and st.row(i)[2].value != 0:
                reTry = st.row(i)[2].value
            mouseClick(2,"left",img,reTry)
            print("双击左键",img)
        elif cmdType.value == 3:
            #3 单击右键
            #取图片名称
            img = st.row(i)[1].value
            reTry = 1
            if st.row(i)[2].ctype == 2 and st.row(i)[2].value != 0:
                reTry = st.row(i)[2].value
            mouseClick(1,"right",img,reTry)
            print("单击右键",img)
        elif cmdType.value == 4:
            #4 输入
            inputValue = st.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
            print(f"输入{inputValue}")
        elif cmdType.value == 5:
            #5 等待
            waitTime = st.row(i)[1].value
            time.sleep(waitTime)
            print(f"等待{waitTime}s")
        elif cmdType.value == 6:
            #6 滚轮
            scrol1 = st.row(i)[1].value
            pyautogui.scroll(int(scrol1))
            print(f"滚轮欢动{scrol1}距离")
        i+=1

if __name__ == "__main__":
    file = f'{Dir}\cmd.xls'
    #打开文件
    wb = xlrd.open_workbook(filename=file)
    #获取表格sheet页
    st = wb.sheet_by_index(0)

    # max_raw = list(st.dimensions)[-1]
    print("欢迎使用自助行RPA~~~")
    if checkCmd :=dataCheck(st):
        key = input('选择功能： 1.做一次 2.做到死 \n')
        if key == "1":
            #循环获取指令
            MainWork(st)
        elif key == '2':
            MainWork(st)
            time.sleep(0.1)
            print("等待0.1秒")
    else:
        print('输入有误或已退出！')
