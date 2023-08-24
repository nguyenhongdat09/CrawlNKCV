import pandas
import pandas as pd
from Constant.Constant import Constant as ct
import numpy as np
import sys
pd.set_option('display.max_columns', None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('display.max_rows', None)

import datetime
class readExcel:
    def __init__(self):
        self.ten_excel = ct.ten_file_excel
    def hour_calculate(self):
        tu_ngay, den_ngay = datetime.datetime.strptime(ct.tu_ngay, '%d/%m/%Y'), datetime.datetime.strptime(
            ct.den_ngay, '%d/%m/%Y')
        df = pd.DataFrame(columns=['ngay', 'gio'])
        while tu_ngay <= den_ngay:
            if tu_ngay.weekday() >= 0 and tu_ngay.weekday() <= 5:
                dict = {
                    "ngay": tu_ngay
                }
                if tu_ngay.weekday() == 5:
                    dict['gio'] = 4
                else:
                    dict['gio'] = 8
                temp_df = pd.DataFrame([dict])
                df = pd.concat([df, temp_df], ignore_index=True)
            tu_ngay += datetime.timedelta(days=1)
        return df
    def read_excel(self, sheet_name):
        df = pd.read_excel(ct.ten_file_excel, sheet_name = sheet_name, usecols= ["stt_rec", "ma_da", "noi_dung", "ngay_ht", "so_gio", "diem"])
        df["ngay_ht"] = pd.to_datetime(df["ngay_ht"], format="%d/%m/%Y")
        df['so_gio'] =  df['so_gio'].fillna(0)
        row_empty_hour:pandas.DataFrame = df.loc[df['so_gio'] <= 0]
        if not row_empty_hour.empty:
            print(row_empty_hour)
            a = input("Các dòng trên chưa nhập giờ hãy xem lại. Có muốn vẫn tiếp tục không (1: Có, 0: Không): ")
            while a not in ('0', '1'):
                a = input("Nhập chưa đúng. Có muốn vẫn tiếp tục không (1: Có, 0: Không): ")
            if a == '1':
                return df
            else:
                sys.exit()

    def update_nkcv_da_nhap(self, df_nkcv, df_ngay):
        df_loc = df_nkcv.loc[df_nkcv['stt_rec'].str.contains('NK1') > 0, ['ngay_ht', 'so_gio']]
        merged_df = df_ngay.merge(df_loc, how='left', left_on='ngay', right_on='ngay_ht')
        merged_df['gio'] = merged_df['gio'] - merged_df['so_gio'].fillna(0)
        merged_df = merged_df.drop(columns=['ngay_ht', 'so_gio'])
        return merged_df
    def allocate_nkcv(self) -> pd.DataFrame:
        df_nkcv = self.read_excel("Nhat ky chua nhap")
        df_ngay = self.hour_calculate()
        df_ngay = self.update_nkcv_da_nhap(df_nkcv, df_ngay)
        df_ngay = df_ngay.loc[df_ngay['gio'] > 0]
        df_ngay['gio_phan_bo'], df_ngay['stt_rec'], df_ngay['noi_dung'], df_ngay['gio_pb_arr'], df_ngay['diem']= 0, '', '', '', ''
        df_nk = df_nkcv.loc[df_nkcv['stt_rec'].str.contains('NK1') > 0]
        df_yc = df_nkcv.loc[df_nkcv['stt_rec'].str.contains('YC1') > 0]
        result = pd.merge(df_nk, df_yc, on='noi_dung', how='inner')
        rows_to_drop = df_yc[df_yc['noi_dung'].isin(result['noi_dung'])].index
        df_nkcv = df_yc.drop(rows_to_drop)
        df_nkcv['diem'] = df_nkcv['diem'].fillna(0)
        for index, row in df_nkcv.iterrows():
            stt_rec, so_gio, diem, noi_dung = row['stt_rec'], row['so_gio'], row['diem'], row['noi_dung']
            while so_gio > 0 :
                df = df_ngay.loc[df_ngay['gio_phan_bo'] < df_ngay['gio']].head(1)
                if df.empty:
                    break
                gio_phan_bo = np.where((df['gio'] - df['gio_phan_bo'] ) < so_gio, (df['gio'] - df['gio_phan_bo'] ),  so_gio )[0]
                so_gio = so_gio - np.where((df['gio'] - df['gio_phan_bo'] ) < so_gio, (df['gio'] - df['gio_phan_bo'] ),  so_gio )[0]
                df_ngay.loc[df_ngay['ngay'] == df['ngay'].values[0], 'gio_phan_bo'] = df['gio_phan_bo'] + gio_phan_bo
                df_ngay.loc[df_ngay['ngay'] == df['ngay'].values[0], 'stt_rec'] +=   ',' + stt_rec
                df_ngay.loc[df_ngay['ngay'] == df['ngay'].values[0], 'noi_dung'] += ',' + noi_dung
                df_ngay.loc[df_ngay['ngay'] == df['ngay'].values[0], 'diem'] += ',' + str(diem)
                df_ngay.loc[df_ngay['ngay'] == df['ngay'].values[0], 'gio_pb_arr'] += ',' + f'{gio_phan_bo}'
        return df_ngay


    def set_diem(self, row):
        ma_diem_dict = dict(ct.ma_diem_tuple)
        diem = float(row['diem'])
        diem_int = int(diem)

        ma_diem = ma_diem_dict.get(str(diem_int))
        if ma_diem is not None:
            return ma_diem
        else:
            return '05'
    def get_df_nkcv_finish(self):
        df = self.allocate_nkcv()
        df = df.loc[df['stt_rec'] != '']
        df = df.explode('stt_rec', ignore_index=True)
        df['stt_rec'] = df['stt_rec'].str.split(',').str[1:]
        df['gio_pb_arr'] = df['gio_pb_arr'].str.split(',').str[1:]
        df['diem'] = df['diem'].str.split(',').str[1:]
        df['noi_dung'] = df['noi_dung'].str.split(',').str[1:]
        # Sử dụng explode để tách thành nhiều dòng
        df = df.explode('stt_rec', ignore_index=True)
        # Đánh stt
        df["stt"] = df.groupby("ngay")["ngay"].rank(method="first", ascending=True)
        df['gio_pb_arr'] = df.apply(lambda row: row['gio_pb_arr'][int(row['stt']-1)], axis=1)
        df['diem'] = df.apply(lambda row: row['diem'][int(row['stt'] - 1)], axis=1)
        df['noi_dung'] = df.apply(lambda row: row['noi_dung'][int(row['stt'] - 1)], axis=1)
        df["stt"] = df.groupby("stt_rec")["stt_rec"].rank(method="first", ascending=True)
        df_group_by_stt_rec = df.groupby("stt_rec")['stt'].max().reset_index()
        df_group_by_stt_rec['trang_thai'] = 'HT'
        df = df.merge(df_group_by_stt_rec, how='left', left_on=['stt_rec','stt'], right_on=['stt_rec','stt'])
        df['trang_thai'] = df['trang_thai'].fillna('TH')
        df.loc[df['trang_thai'] == 'TH', 'diem'] = '0'
        df['ma_diem'] = df.apply(lambda row: self.set_diem(row), axis=1)
        df['nh_cv1'], df['nh_cv2'], df['ma_nv'] = ct.nh_cv1, ct.nh_cv2, ct.user_name
        df.drop('gio_phan_bo', axis=1, inplace=True)
        df.drop('stt', axis=1, inplace=True)
        return df
# reader = readExcel()
# reader.get_df_nkcv_finish()
