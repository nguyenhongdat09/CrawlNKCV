import os
import json
class Constant:
    path = os.getcwd() + '\Infomation.js'
    data = ''
    with open(path, 'r') as f:
        data = json.load(f)
    user_name = data['username']
    password = data['password']
    url = data['url']
    url_nkcv = data['url_nkcv']
    url_qlyc = data['url_qlyc']
    tu_ngay = data['tu_ngay']
    den_ngay = data['den_ngay']
    trang_thai = data['trang_thai']
    ten_file_excel = data['ten_file_excel']
    ma_diem_tuple = [('0', '05'), ('1', '01'), ('2', '02'), ('3', '03'), ('4', '04'), ('5', '05')]
    nh_cv1 = data['nh_cv1']
    nh_cv2 = data['nh_cv2']
