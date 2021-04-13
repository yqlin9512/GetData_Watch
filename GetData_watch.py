#encoding: utf-8
# import requests
from requests import get
from re import findall
from time import sleep
import pandas as pd
from bs4 import BeautifulSoup
from tqdm import tqdm
import os
import tkinter as tk
from tkinter import filedialog




# 定义参数
headers = {'Host':'www.xbiao.com',
           'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.82 Safari/537.36 Edg/89.0.774.50',
           'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
           'Accept-Language':'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
           'Accept-Encoding':'gzip, deflate',
           'Connection':'keep-alive'}
list_url = 'http://www.xbiao.com/search/index?tp=p&wd=%(key)s'          #腕表之家，以货号查询
#list_url = 'http://s.taobao.com/search?q=%(key)s&ie=utf8&filter_tianmao=tmall&s=%(page)d&filter=%(fprice)s&auction_tag[]=%(new)d'
#cookies ='bdshare_firstime=1614857012880; block=å®‡èˆ¶$http://bbs.xbiao.com/hublot/|æ¬§ç±³èŒ„$http://bbs.xbiao.com/omega/; BIGipServertop-xbiao-web=4274523914.20480.0000; Hm_lvt_a614a6a498e45b564870acd236f789d4=1614912452,1614912520,1615360218,1615538221; PHPSESSID=meua750cli17bnt7gfgc82rls5; Hm_lpvt_a614a6a498e45b564870acd236f789d4=1615540393'

# 正则模式
p_img = '<img src="(.*)" alt'         #图片URL


def Get_watchinfo(dirpath,filename):
    fpath = os.path.join(dirpath, filename)
    watch = pd.read_excel(fpath)
    try:
        pcode = watch['厂商货号']
        savebasic = filename.split('.')[0] + "_basic.csv"
        saveparam = filename.split('.')[0] + "_param.csv"

        col_basic = ['图片链接', '商品名称', '系列', '款式', '材质', '价格', '商品网站链接', '喜欢度']
        col_param = ['编号', '品牌', '系列', '机芯类型', '性别',
                     '机芯型号', '机芯类型', '机芯直径', '机芯厚度', '振频', '宝石数', '电池寿命', '动力储备', '技术认证',
                     '表径', '表壳厚度', '表盘颜色', '表盘形状', '表带颜色', '表扣类型', '背透', '重量', '防水深度', '表扣间距', '表耳间距',
                     '表壳材质', '表盘材质', '表镜材质', '表冠材质', '表带材质', '表扣材质', '其他功能']
        watch_basic = {}
        watch_data = {}
        # p_basic #['照片路径'，'商品名称', '基础系列：', '款式：（机芯、size、男女）', '基础材质：', '价格：','商品详情url','喜欢']
        # 数据爬取

        for i in tqdm(range(pcode.shape[0])):
            pkey = pcode[i].strip("\t")
            print("开始查询：", pkey)
            dic_basic = dict.fromkeys(col_basic, '-')
            dic_param = dict.fromkeys(col_param, '-')
            try:
                # 进入查询主界面爬取基础信息
                url = list_url % {'key': pkey}
                res = get(url, headers=headers)
                res.encoding = 'utf-8'
                html = res.text
                # print(html)
                img = findall(p_img, html)
                soup = BeautifulSoup(html, 'lxml')
                p_items = []
                for i in soup.find_all('ul', class_="s_attr"):
                    for j in i.find_all('li'):
                        p_items.append(j.get_text())
                purl = p_items[-1]
                purl_param = purl + "param.html"
                dic_basic['图片链接'] = img[0]
                dic_basic['商品名称'] = p_items[0]
                for i1 in range(1, len(p_items) - 1):
                    pitem = p_items[i1].split('：')
                    if len(pitem) == 2:
                        if pitem[0] == '系列':
                            dic_basic['系列'] = pitem[1]
                        elif pitem[0] == '款式':
                            dic_basic['款式'] = pitem[1]
                        elif pitem[0] == '材质':
                            dic_basic['材质'] = pitem[1]
                        elif pitem[0] == '价格':
                            dic_basic['价格'] = pitem[1]
                dic_basic['商品网站链接'] = purl
                # 进入商品详情url爬取喜欢数量
                res_p = get(purl, headers=headers)
                res_p.encoding = 'utf-8'
                html_p = res_p.text
                soup_p = BeautifulSoup(html_p, 'lxml')
                for li in soup_p.find_all('div', class_="handle_btn clearfix"):
                    like = li.get_text().split('\n')[1]
                    dic_basic['喜欢度'] = like
                watch_basic[pkey] = dic_basic
                # print(type(keydic), type(dic), dic['商品名称'], keydic[pkey])
                # 进入商品的详细参数url爬取详细基础信息
                res_param = get(purl_param, headers=headers)
                res_param.encoding = 'utf-8'
                html_param = res_param.text
                soup_param = BeautifulSoup(html_param, 'lxml')
                p_param = []
                for ii in soup_param.find_all('td', class_="param_info_txt"):
                    for jj in ii.find_all('li'):
                        p_param.append(jj.get_text())
                # print('p_param', len(p_param), p_param)
                param = []
                fuc_num = 0
                for k in range(len(p_param)):
                    pt = p_param[k].split('：')
                    if len(pt) == 2:
                        if pt[0] in dic_param.keys():
                            dic_param[pt[0]] = pt[1]
                    else:
                        pt = p_param[k].split('材质')
                        if len(pt) == 2:
                            pt_name = pt[0] + '材质'
                            if pt_name in dic_param.keys():
                                dic_param[pt_name] = pt[1]
                        else:
                            fuc_num += 1
                            if fuc_num == 1:
                                dic_param['其他功能'] = p_param[k]
                            else:
                                dic_param['其他功能'] += '&' + p_param[k]
                watch_data[pkey] = dic_param
                print(pkey, '查询成功！')
                sleep(3)
            except:
                print(pkey, "查询失败，继续")
                continue

        res_basic = []
        res_param = []
        for key, value in watch_data.items():
            res_param.append(value)
        for key, value in watch_basic.items():
            res_basic.append(value)
        pd_param = pd.DataFrame(res_param, columns=col_param)
        pd_basic = pd.DataFrame(res_basic, columns=col_basic)
        pd_basic.fillna(' ', inplace=True)
        pd_param.fillna(' ', inplace=True)
        pd_param.to_csv(os.path.join(dirpath, saveparam), encoding='utf-8_sig', index=False)
        pd_basic.to_csv(os.path.join(dirpath, savebasic), encoding='utf-8_sig', index=False)
        print("==========", filename, '保存成功！==========')
    except:
        print('请确保手表的查询货号标题为 ”厂商货号”')


if __name__ == '__main__':
    print("==========使用事项==========")
    print("1.选取文件夹，手表文件需为xlsx后缀名")
    print("2.手表文件内仅支持1个sheet,第一行为标题，其中查询的手表货号标题需为“厂商货号”")
    #打开选择文件夹对话框
    root = tk.Tk()
    root.withdraw()
    dirpath = filedialog.askdirectory()     #获取选择好的文件夹
    # dirpath = r"C:\Users\HTDF\Desktop\PY\watch\lyq"
    filelist = []
    for filename in os.listdir(dirpath):
        if os.path.splitext(filename)[1] == '.xlsx':
            filelist.append(filename)
    for fn in filelist:
        print('==========开始处理：',fn,"==========")
        Get_watchinfo(dirpath,fn)