import requests,openpyxl,os
has = os.path.exists('music.xlsx')
if has:
    wb = openpyxl.load_workbook('music.xlsx')
    sheet = wb['music']
else:
    wb=openpyxl.Workbook()  
    #创建工作簿
    sheet=wb.active 
    #获取工作簿的活动表
    sheet.title='music' 
    #工作表重命名

    sheet['A1'] ='歌曲名'     #加表头，给A1单元格赋值
    sheet['B1'] ='歌手'
    sheet['C1'] ='所属专辑'   #加表头，给B1单元格赋值
    sheet['D1'] ='播放时长'   #加表头，给C1单元格赋值
    sheet['E1'] ='歌曲链接'   #加表头，给C1单元格赋值
    sheet['F1'] ='完整歌词'   #加表头，给D1单元格赋值


url = 'https://c.y.qq.com/soso/fcgi-bin/client_search_cp'
for x in range(500):
    # params = {
    #     'ct': '24',
    #     'qqmusic_ver': '1298',
    #     'new_json': '1',
    #     'remoteplace': 'txt.yqq.song',
    #     'searchid': '64405487069162918',
    #     't': '0',
    #     'aggr': '1',
    #     'cr': '1',
    #     'catZhida': '1',
    #     'lossless': '0',
    #     'flag_qc': '0',
    #     'p': str(x + 1),
    #     'n': '20',
    #     'w': '华语',
    #     'g_tk': '5381',
    #     'loginUin': '0',
    #     'hostUin': '0',
    #     'format': 'json',
    #     'inCharset': 'utf8',
    #     'outCharset': 'utf-8',
    #     'notice': '0',
    #     'platform': 'yqq.json',
    #     'needNewCode': '0'
    # }
    params = {
        'ct': '24',
        'qqmusic_ver': '1298',
        'remoteplace': 'txt.yqq.center',
        'searchid': '49447739519928628',
        'aggr': '0',
        'catZhida': '1',
        'lossless': '0',
        'sem': '1',
        't': '7',
        'p': str(x + 1),
        'n': '5',
        'w': '粤语老歌',
        'g_tk_new_20200303': '5381',
        'g_tk': '5381',
        'loginUin': '0',
        'hostUin': '0',
        'format': 'json',
        'inCharset': 'utf8',
        'outCharset': 'utf-8',
        'notice': '0',
        'platform': 'yqq.json',
        'needNewCode': '0'
    }

    res_music = requests.get(url, params=params)
    json_music = res_music.json()
    list_music = json_music['data']['lyric']['list']
    print(list_music)
    for music in list_music:
        name = music['songname']
        # 以name为键，查找歌曲名，把歌曲名赋值给name
        album = music['albumname']
        if len(music['singer']) > 1:
            singer = music['singer'][0]['name'] + '/' + music['singer'][1]['name']
        else:
            singer = music['singer'][0]['name']
        # 查找专辑名，把专辑名赋给album
        time = music['interval']
        # 查找播放时长，把时长赋值给time
        link = 'https://y.qq.com/n/yqq/song/' + str(music['media_mid']) + '.html\n\n'
        # 查找播放链接，把链接赋值给link
        lyrics = music['content']

        sheet.append([name,singer,album,time,link,lyrics])
        # 把name、album、time和link写成列表，用append函数多行写入Excel
        # print('歌曲名：' + name + '\n' + '歌手:' + singer +'\n' + '所属专辑:' + album +'\n' + '播放时长:' + str(time) + '\n' + '播放链接:'+ link + '\n' + '完整歌词:'+ lyrics)
        
wb.save('music.xlsx')            
# #最后保存并命名这个Excel文件