# coding=utf-8
import requests
import json
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import shutil
import sys
import os
import uuid

reload(sys)
sys.setdefaultencoding('utf8')

album_list = []


def query_category(code, sub_code, pageSize):
    url = 'https://m.ximalaya.com/m-revision/page/category/queryCategoryPage?categoryCode={}&subCategoryCode={}&pageSize={}'.format(
        code, sub_code, pageSize)
    r = requests.get(url)
    return r.text


def get_image(file_name):
    try:
        if file_name.find('http') == -1:
            url = 'https://imagev2.xmcdn.com/' + file_name + "!op_type=3&columns=144&rows=144"
        else:
            url = file_name

        res = requests.get(url)
        file_name = 'tmp/' + str(uuid.uuid4()) + '.jpg'
        with open(file_name, 'wb') as f:
            f.write(res.content)
        f.close()
    except Exception as e:
        print 'get_image_Exception'
    return file_name


def get_key(obj, key, sub_key):
    return obj[key].get(sub_key, '')


def to_execl():
    wb = Workbook()
    ws = wb.active

    head = ['ID', '专辑标题', '子标题', '今日播放量', '总播放量', '图标', '用户昵称', '头像', '个性签名']
    ws.append(head)
    row = 1
    for album in album_list:
        print '处理了{}行'.format(row)
        row += 1
        ws.row_dimensions[row].height = 80
        try:
            info = [
                album['id'],
                get_key(album, 'albumInfo', 'title'),
                get_key(album, 'albumInfo', 'customTitle'),
                get_key(album, 'statCountInfo', 'trackCount'),
                get_key(album, 'statCountInfo', 'playCount'),
                "",
                get_key(album, 'anchorInfo', 'nickname'),
                "",
                get_key(album, 'anchorInfo', 'personalSignature'),
            ]
            ws.append(info)

            # 专辑头
            cover_url = get_key(album, 'albumInfo', 'cover')
            if cover_url != '':
                cover_img = Image(get_image(cover_url))
                cover_img.height = 100
                cover_img.width = 100
                ws.add_image(cover_img, 'F' + str(row))

            # 用户头像
            logo_url = get_key(album, 'anchorInfo', 'logo')
            if logo_url != '':
                user_logo = Image(get_image(logo_url))
                user_logo.height = 100
                user_logo.width = 100
                ws.add_image(user_logo, 'H' + str(row))
        except Exception as e:
            print album['id'] + '----Exception'
    wb.save('xmly.xlsx')


if __name__ == '__main__':
    os.makedirs('./tmp')
    album_text = query_category('qinggan', 'qinggan', 1000)
    album_data = json.loads(album_text)
    album_list = album_data['data']['firstPageCategoryAlbums']['albumBriefDetailInfos']
    to_execl()
    shutil.rmtree("tmp")
    print 'xmly ok > xmly.xlsx'
