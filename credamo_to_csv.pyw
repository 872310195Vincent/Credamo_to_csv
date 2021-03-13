from json import loads
import csv
from urllib import request
from datetime import datetime
from os import remove
import tkinter as tk
import platform

if platform.platform().startswith('macOS'):
	import ssl
	ssl._create_default_https_context = ssl._create_unverified_context


def get_user_info_gui(data_request_header):
    # 1.从界面输入读取用户信息，保存user_info列表，里面包含url变量和headers字典
    header_components = data_request_header.split('\n')
    # 储存url, headers(包含User-Agent, Referer和Cookie信息)
    header_infos = ["User-Agent", "Referer", "Cookie"]
    headers = {}
    url = ''
    for header_component in header_components:
        if 'currPageIndex=' in header_component:
            end_loc = header_component.rfind('currPageIndex=') + len('currPageIndex=')
            start_loc = header_component.find(' ')
            # 需要在url最后加上数字页码
            url = "https://www.credamo.com" + header_component[start_loc + 1:end_loc]
        for header_info in header_infos:
            if header_component.startswith(header_info.lower()) or header_component.startswith(header_info):
                headers[header_info] = header_component[len(header_info) + 2:].rstrip()
    user_info = [url, headers]
    return user_info


# 以下是测试get_user_info_gui()的代码，已对Safari，Edge，Chrome，Firefox进行测试
# with open('Safari请求头') as file_object:
#     request_header = file_object.read()
# print(get_user_info_gui(request_header))

def get_data_json(user_info, page):
    # 2.获取json格式数据
    url = user_info[0]
    headers = user_info[1]
    req = request.Request(url + str(page), headers=headers)  # 添加包含Credamo用户信息的头部信息
    with request.urlopen(req) as res:
        return loads(res.read().decode('utf-8-sig'))


def get_vari_info(jsonfile):
    # 3.解析json数据，返回vari_info = [headers, question_names, q_nums]信息
    if jsonfile['success']:
        loaded_file = jsonfile['data']
        headers_dic = loaded_file['header']
        # 处理变量信息，将变量id储存为头部信息headers，将变量名和变量标签储存为前两行
        headers = ['userId', 'answerId', 'status']
        question_names = ['userId', 'answerId', 'status']
        q_nums = ['userId', 'answerId', 'status']
        for header in headers_dic:
            q_id = header['id']
            headers.append(q_id)
            question_names.append(header['questionName'])
            q_nums.append(header['qNum'])
        vari_info = [headers, question_names, q_nums]
        return vari_info
    else:
        return []


def write_vari_info(vari_info, filename):
    # 4.打开第一个文件时，需要写入变量名
    with open(filename, 'w', newline='', encoding='UTF-8-sig')as f:
        f_csv = csv.writer(f)
        # 写入变量id
        f_csv.writerow(vari_info[0])
        # 写入变量内容，编号
        f_csv.writerows(vari_info[1:])


def get_vari_value(jsonfile):
    # 5.解析json数据，返回vari_info = [headers, question_names, q_nums]信息
    if jsonfile['success']:
        loaded_file = jsonfile['data']
        headers_dic = loaded_file['header']
        rows = loaded_file['rowList']
        # 处理变量信息，将变量id储存为头部信息headers，将变量名和变量标签储存为前两行
        headers = ['userId', 'answerId', 'status']
        question_names = ['userId', 'answerId', 'status']
        q_nums = ['userId', 'answerId', 'status']
        for header in headers_dic:
            q_id = header['id']
            headers.append(q_id)
            question_names.append(header['questionName'])
            q_nums.append(header['qNum'])
        return headers, rows
    # else:
    #     print("下载代码失败")


# def write_vari_value(headers, rows, filename):
#     # 6.写入值
#     with open(filename, 'a', newline='', encoding='UTF-8-sig')as f:
#         f_csv_dic = csv.DictWriter(f, headers)
#         f_csv_dic.writerows(rows)

def write_vari_value_accepted(headers, rows, filename, statuses):
    # 6.写入值
    with open(filename, 'a', newline='', encoding='UTF-8-sig')as f:
        f_csv_dic = csv.DictWriter(f, headers)
        rows_accepted = []
        for row in rows:
            # 判断状态，尚未采纳或拒绝1，模拟 2，采纳 3，手动拒绝 5，自动拒绝 6
            for status in statuses:
                if row['status'] == status:
                    rows_accepted.append(row)
        f_csv_dic.writerows(rows_accepted)


def delete_row_columns(input_file, output_file):
    # 7.删除多余的前3列，第1行
    with open(input_file, newline='', encoding='utf-8-sig') as f_in, \
            open(output_file, 'a', newline='', encoding='utf-8-sig') as f_out:
        reader = csv.reader(f_in)
        writer = csv.writer(f_out)
        # 跳过第一行
        next(reader)
        for row in reader:
            writer.writerow(row[3:])


# def main_func(max_page_num,
#               request_header = '请求头.txt',
#               save_file = 'credamo_data',
#               add_time = True):
#     '''这是主程序，运行前面所有函数'''
#     save_file += 'temp.csv'
#     # 请复制请求头到txt文件
#     user_info = get_user_info_txt(request_header)
#     # 读取第一个文件
#     jsonfile = get_data_json(user_info, 1)
#     # 得到并写入变量信息
#     vari_info = get_vari_info(jsonfile)
#     write_vari_info(vari_info, save_file)
#     print('成功写入变量信息')
#     # 写入多个页面的值
#     for page in range(max_page_num):
#         page += 1
#         current_jsonfile = get_data_json(user_info, page)
#         headers, rows = get_vari_value(current_jsonfile)
#         write_vari_value_accepted(headers, rows, save_file)
#         print('成功写入第'+str(page)+'页变量值')
#     # 删除前3列和第一行
#     # 判断是否添加时间
#     if add_time:
#         # 获取下载文件时间
#         make_file_time = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
#         save_file_formatted = save_file[:-8] + "_" + make_file_time +'.csv'
#         print('成功添加时间')
#     else:
#         save_file_formatted = save_file[:-8] + '.csv'
#     delete_row_columns(save_file, save_file_formatted)
#     print('成功删除多余行列')
#     # 最终删除临时文件
#     os.remove(save_file)
#     print('成功删除临时文件')
'''
max_page_num = 4, 最大页数
request_header = '请求头.txt', 请将
save_file = 'credamo_data', 文件名
add_time = True，是否在文件名后面加时间
'''
# main_func(max_page_num = 4, request_header = '请求头.txt', save_file = 'credamo_data', add_time = True)

'''创建GUI界面'''
# 1.创建主页面
window = tk.Tk()
window.title("下载Credamo数据")
window.geometry("800x1080")

# 2.创建一个文本展示框，可用于展示程序运行结果，输入文字
request_header_instruc = tk.Label(window, text="1.请在下方框中粘贴请求头：",
                                  # bg="green", font=('Arial',12),
                                  width=40, height=1)
request_header_instruc.pack()
request_header = tk.Text(window, height=20)
request_header.pack()

# 3.创建一个输入框，用于输入页码
# page_instruc = tk.Label(window, text="2.请在下方框中输入网页中显示的最大页码：",
#                         # bg="green", font=('Arial',12),
#                         width=40, height=1)
# page_instruc.pack()
# get_page = tk.Text(window, width=4, height=1)
# get_page.insert('0.0', '1')
# get_page.pack()

# 3.创建一个输入框，用于输入文件名
if platform.platform().startswith('macOS'):
	name_instruc = tk.Label(window, text="2.请输入下载数据的地址\文件名：",
                        # bg="green", font=('Arial',12),
                        width=40, height=1)
else:
	name_instruc = tk.Label(window, text="2.请输入下载数据的文件名：",
                        # bg="green", font=('Arial',12),
                        width=40, height=1)
name_instruc.pack()
get_name = tk.Text(window, width=30, height=1)
if platform.platform().startswith('macOS'):
	get_name.insert('0.0', 'Desktop\credamo_data')
else:
	get_name.insert('0.0', 'credamo_data')

get_name.pack()
time_instruc = tk.Label(window, text="3.请勾选是否在文件名后添加时间：",
                        # bg="green", font=('Arial',12),
                        width=40, height=1)
time_instruc.pack()
time_var = tk.BooleanVar()  # 整数变量
time_c = tk.Checkbutton(window, text='是',
                        variable=time_var, onvalue=True, offvalue=False)  # 选定时候值为1
time_c.pack()
time_c.select()


def choose_status():
    # 创建多选框，可以选择下载的被试状态，为未采纳或拒绝，已采纳，或模拟
    statuses = []
    if var1.get() == 1:
        statuses.append(1)
    if var2.get() == 1:
        statuses.append(2)
    if var3.get() == 1:
        statuses.append(3)
    return statuses


status_instruc = tk.Label(window, text="4.请选择要下载被试的种类：",
                          # bg="green", font=('Arial',12),
                          width=40, height=1)
status_instruc.pack()
var1 = tk.IntVar()  # 整数变量
var2 = tk.IntVar()  # 整数变量
var3 = tk.IntVar()  # 整数变量
c1 = tk.Checkbutton(window, text='未采纳或拒绝',
                    variable=var1, onvalue=1, offvalue=0,
                    command=choose_status)  # 选定时候值为1
c2 = tk.Checkbutton(window, text='模拟',
                    variable=var2, onvalue=1, offvalue=0,
                    command=choose_status)  # 选定时候值为1
c3 = tk.Checkbutton(window, text='已采纳',
                    variable=var3, onvalue=1, offvalue=0,
                    command=choose_status)  # 选定时候值为1
c1.pack()
c3.pack()
c2.pack()
c1.select()
c3.select()


# 4.GUI的主程序
def download_data_gui():
    global_check = True
    # 1.获得请求头并检查
    res_header = request_header.get('0.0', 'end')
    user_info = get_user_info_gui(res_header)
    if user_info[0] == '' or ('User-Agent' not in user_info[1].keys()) or ('Referer' not in user_info[1].keys()) or (
            'Cookie' not in user_info[1].keys()):
        show_execute_info('请求头内容有误，请点击’请求头在哪里？‘按钮阅读如何复制请求头信息')
        global_check = False
    # # 2.获取并检查页码
    # max_page_num = get_page.get('0.0', 'end')
    # try:
    #     max_page_num = int(max_page_num)
    #     if max_page_num <= 0:
    #         show_execute_info('请输入正整数页码，谢谢')
    #         global_check = False
    # except ValueError or TypeError:
    #     show_execute_info('请输入正整数页码，谢谢')
    #     global_check = False
    # 获取并检查文件名称
    save_file = 'credamo_data'
    if get_name.get('0.0', 'end'):
        save_file = get_name.get('0.0', 'end')
        save_file = save_file.strip()
    else:
        show_execute_info('您没有指定生成数据文件的名称，将默认生成名为credamo_data的数据')
    if global_check:
        save_file += ' temp.csv'
        # 读取第一个文件
        jsonfile = get_data_json(user_info, 1)
        # 得到并写入变量信息
        vari_info = get_vari_info(jsonfile)
        if vari_info:
            write_vari_info(vari_info, save_file)
            show_execute_info('成功写入变量信息')
        else:
            show_execute_info('获取变量信息失败，请检查输入')
        # 写入多个页面的值
        page = 1
        while True:
            current_jsonfile = get_data_json(user_info, page)
            headers, rows = get_vari_value(current_jsonfile)
            if len(rows) > 0:
                write_vari_value_accepted(headers, rows, save_file, statuses=choose_status())
                show_execute_info('成功写入第' + str(page) + '页变量值')
                page += 1
            else:
                break
        # 是否在文件名称后添加时间
        add_time = time_var.get()
        if add_time:
            # 获取下载文件时间
            make_file_time = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
            save_file_formatted = save_file[:-8] + "_" + make_file_time + '.csv'
            show_execute_info('已在文件名后添加时间')
        else:
            save_file_formatted = save_file[:-8] + '.csv'
        # 删除前3列和第一行
        delete_row_columns(save_file, save_file_formatted)
        # show_execute_info('成功删除多余行列')
        # 最终删除临时文件
        remove(save_file)
        show_execute_info('下载数据完成'+make_file_time)


# 4.1一个按钮，点击后下载数据
download_button = tk.Button(window, text='下载', width=5, height=1,
                            command=download_data_gui)  # 点击按钮运行这个函数
download_button.pack()  # 放到label下面


# 5.创建一个message，动态显示输出
def show_execute_info(var):
    execute_info.insert("0.0", var + '\n')
    # t.insert(1.1, var) #在第一行第一个字后面插入


execute_info = tk.Text(window, height=4)
first_info = '请粘贴请求头，输入数据名称，选择是否在名称后面添加时间，然后点击下载按钮即可下载数据'
show_execute_info(first_info)
execute_info.pack()


# 6.创建一个窗口，用语指导被试如何复制请求头
def help_copy_header():
    window_help_copy_header = tk.Toplevel(window)  # 在主窗口上的第二层窗口
    window_help_copy_header.geometry('600x600')
    window_help_copy_header.title('复制请求头的方法')
    var = tk.StringVar()
    help_content = tk.Message(window_help_copy_header, width=340,
                              text='请选择您使用的浏览器')

    def print_selection():
        help_content.config(text=var.get())  # 用于将选择的var，值为value输出

    chrome_copy_header = 'Chrome或者基于Chromium的浏览器(如Edge)中复制请求头的方法：\n' \
                         '1.在浏览器中打开您要下载数据的项目，打开数据清理页面，您可以浏览您要下载的数据。\n' \
                         '2.在页面空白处点击鼠标右键，再点击”检查“，可以看到浏览器弹出一个调试网页的工具。\n' \
                         '3.在这个工具最顶部一行，点击”网络“，可以看到”记录网络活动，执行一个请求……“\n' \
                         '4.刷新网页，可以看到调试工具不断呈现浏览器的请求，' \
                         '在刷新的内容上方有一条工具栏可以筛选这些请求：ALL XHR JS CSS……，请点击XHR，进行筛选' \
                         '关注请求内容的‘名称’一列，看到最后一行，包含currPageSi……字样。\n' \
                         '5.双击这一行，会弹出头部信息内容，请复制”请求头“Request Header中的内容，' \
                         '从”:authority“开始，到”user-agent:”的全部内容结束\n' \
                         '6.最后选择本软件第一个窗口，在框中粘贴内容即可。'
    firefox_copy_header = 'Firefox浏览器中复制请求头的方法：\n' \
                          '1.在浏览器中打开您要下载数据的项目，打开数据清理页面，您可以浏览您要下载的数据。\n' \
                          '2.在页面空白处点击鼠标右键，再点击”检查元素“，可以看到浏览器弹出一个调试网页的工具。\n' \
                          '3.在这个工具最顶部一行，点击”↑↓网络“，可以看到”请进行至少一项请求……“\n' \
                          '4.刷新网页，可以看到调试工具不断呈现浏览器的请求，请注意‘文件’这一列，滚动鼠标滚轮向下浏览最新请求，' \
                          '直到‘文件’一列看到包含currPageSize=……&currPageIndex=字样的一行。\n' \
                          '5.把光标定位在这一行，点击鼠标右键，选择”复制“，点击”复制请求头“。\n' \
                          '6.最后选择本软件第一个窗口，在框中粘贴内容即可。'
    safari_copy_header = 'Safari中复制请求头的方法：\n' \
                         '1.打开浏览器后，点击左上角Safari浏览器，选择偏好设置，切换到“高级”选项卡，' \
                         '勾选“在菜单栏中显示开发菜单”' \
                         '2.在浏览器中打开您要下载数据的项目，打开数据清理页面，您可以浏览您要下载的数据。\n' \
                         '2.在页面空白处点击鼠标右键，再点击”检查“，可以看到浏览器弹出一个调试网页的工具。\n' \
                         '3.在这个工具顶部一行，点击”网络“，\n' \
                         '4.刷新网页，可以看到调试工具不断呈现浏览器的请求，' \
                         '在刷新的内容上方有一条工具栏可以筛选这些请求：ALL Document CSS …… XHR Other，请点击XHR，进行筛选' \
                         '关注请求内容的‘名称’一列，看到最后一行，包含一串数字，表明你的项目编号。\n' \
                         '5.双击这一行，会弹出头部信息内容，请复制”请求“中的内容，' \
                         '从”:method:GET“开始，到”Connection:keep-alive”结束\n' \
                         '6.最后选择本软件第一个窗口，在框中粘贴内容即可。'
    r1 = tk.Radiobutton(window_help_copy_header, text='Chrome OR Edge(Chromium-based)', variable=var,
                        value=chrome_copy_header, command=print_selection)
    r2 = tk.Radiobutton(window_help_copy_header, text='Firefox', variable=var,
                        value=firefox_copy_header, command=print_selection)
    r3 = tk.Radiobutton(window_help_copy_header, text='Safari', variable=var,
                        value=safari_copy_header, command=print_selection)
    r1.pack()
    r2.pack()
    r3.pack()
    help_content.pack()


# 6.1创建一个按钮，用于打开如何复制请求头窗口
button_help_copy_header = tk.Button(window, text='请求头在哪里？',
                                    command=help_copy_header).pack()
window.mainloop()
