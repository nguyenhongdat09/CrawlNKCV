import pandas as pd
from Constant.Constant import Constant as ct
import numpy as np
pd.set_option('display.max_columns', None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_colwidth', None)
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
        return df
    def update_nkcv_da_nhap(self, df_nkcv, df_ngay):
        df_loc = df_nkcv.loc[df_nkcv['stt_rec'].str.contains('NK1') > 0, ['ngay_ht', 'so_gio']]
        merged_df = df_ngay.merge(df_loc, how='left', left_on='ngay', right_on='ngay_ht')
        merged_df['gio'] = merged_df['gio'] - merged_df['so_gio'].fillna(0)
        merged_df = merged_df.drop(columns=['ngay_ht', 'so_gio'])
        return merged_df
    def allocate_nkcv(self):
        df_nkcv = reader.read_excel("Nhat ky chua nhap")
        df_ngay = reader.hour_calculate()
        df_ngay = self.update_nkcv_da_nhap(df_nkcv, df_ngay)
        df_ngay['gio_phan_bo'] = []
        df_ngay['stt_rec'] = ''
        df_nkcv = df_nkcv.loc[df_nkcv['stt_rec'].str.contains('YC1') > 0]
        df_nkcv['diem'] = df_nkcv['diem'].fillna(0)
        for index, row in df_nkcv.iterrows():
            stt_rec, so_gio = row['stt_rec'], row['so_gio']
            df_combine = pd.DataFrame(columns=['ngay', 'gio', 'gio_phan_bo', 'stt_rec'])
            while so_gio > 0 :
                df = df_ngay.loc[df_ngay['gio_phan_bo'] < df_ngay['gio']].head(1)
                if df.empty:
                    print(stt_rec)
                    break

                new_val = {
                    'ngay': df['ngay'].values[0],
                    'gio_phan_bo': np.where((df['gio'] - df['gio_phan_bo'] ) < so_gio, (df['gio'] - df['gio_phan_bo'] ),  so_gio )[0]
                }
                so_gio = so_gio - np.where((df['gio'] - df['gio_phan_bo'] ) < so_gio, (df['gio'] - df['gio_phan_bo'] ),  so_gio )[0]
                df_ngay.loc[df_ngay['ngay'] == df['ngay'].values[0], 'gio_phan_bo'] = sum(df_ngay['gio_phan_bo']) + new_val['gio_phan_bo']
                df_ngay.loc[df_ngay['ngay'] == df['ngay'].values[0], 'stt_rec'] = df_ngay['stt_rec'] + ',' + stt_rec
                if df['ngay'].values[0] == '2023-07-22':
                    print(stt_rec)
        print(df_ngay)



        for index, row in df_ngay.iterrows():
            ngay, gio = row['ngay'], row['gio']




reader = readExcel()

reader.allocate_nkcv()
