import os
from typing import Generator, Union
from itertools import chain
import pandas as pd
import openpyxl
from openpyxl.workbook.workbook import Workbook
from deta import Deta


deta = Deta('c09hsnq1_gtaAivrv3sAy4VaqXbo1L4mYssn13SGu')
db = deta.Base('praktikum')


MODUL = {
    # '1': ['E', 'G', 'H', 'I', 'J'],
    '2': ['K', 'M', 'N', 'O', 'P'],
    '3': ['Q', 'S', 'T', 'U', 'V'],
    '4': ['W', 'Y', 'Z', 'AA', 'AB'],
    '5': ['AC', 'AE', 'AF', 'AG', 'AH'],
    '6': ['AI', 'AK', 'AL', 'AM', 'AN'],
    '7': ['AO', 'AQ', 'AR', 'AS', 'AT'],
    '8': ['AU', 'AW', 'AX', 'AY', 'AZ'],
    '9': ['BA', 'BC', 'BD', 'BE', 'BF'],
    '10': ['BG', 'BI', 'BJ', 'BK', 'BL'],
    '11': ['BM', 'BO', 'BP', 'BQ', 'BR'],
    '12': ['BS', 'BU', 'BV', 'BW', 'BX'],
    '13': ['BY', 'CA', 'CB', 'CC', 'CD'],
    '14': ['CE', 'CG', 'CH', 'CI', 'CJ'],
    '15': ['CK', 'CM', 'CN', 'CO', 'CP'],
    '16': ['CQ', 'CS', 'CT', 'CU', 'CV'],
}


class Praktikum:
    def __init__(self, filename: str, matkul: str) -> None:
        wb = openpyxl.load_workbook(filename, data_only=True)
        self.wb = wb
        self.matkul = matkul
        self.kelas = [x for x in wb.sheetnames if wb[x]['E1'].value == 'MODUL 1']
    
    def get_mhs_kelas(self, kelas: str) -> Generator[tuple, None, None]:
        start = 5
        res = []
        while True:
            data = (self.wb[kelas][f'B{start}'].value, self.wb[kelas][f'C{start}'].value)
            if not all(data):
                break
            res.append(data)
            start += 1
        return res
    
    def get_mhs(self, kelas: str = None, as_df: bool = False) -> Union[Generator[tuple, None, None], pd.DataFrame]:
        if not as_df:
            if kelas:
                return self.get_mhs_kelas(kelas)
            else:
                res = []
                for sheet in self.kelas:
                    res.append(self.get_mhs_kelas(sheet))
                return sum(res, [])
        else:
            if kelas:
                nim, nama = zip(*self.get_mhs_kelas(kelas))
            else:
                res = []
                for sheet in self.kelas:
                    res.append(self.get_mhs_kelas(sheet))
                res = sum(res, [])
                nim, nama = zip(*list(res))
            data = pd.DataFrame({
                'nim': nim,
                'nama': nama
            })
            
            data = data.astype({'nim': str})
            return data
    
    def get_per_modul(self, sheet: str, nim: str):
        start = 5
        while True:
            if self.wb[sheet][f'B{start}'].value == nim:
                nama = self.wb[sheet][f'C{start}'].value
                break
            start += 1
        
        kelas = []
        mod = []
        list_nim = []
        praktikan = []
        aktif = []
        kehadiran = []
        awal = []
        jurnal = []
        akhir = []
        total = []
        for i in range(len(MODUL.values())):
            kelas.append(sheet)
            mod.append(i)
            list_nim.append(str(nim))
            praktikan.append(nama)
            hadir = self.wb[sheet][f'{list(MODUL.values())[i][0]}{start}'].value
            aktif.append(True if hadir else False)
            kehadiran.append(hadir)
            awal.append(self.wb[sheet][f'{list(MODUL.values())[i][1]}{start}'].value)
            jurnal.append(self.wb[sheet][f'{list(MODUL.values())[i][2]}{start}'].value)
            akhir.append(self.wb[sheet][f'{list(MODUL.values())[i][3]}{start}'].value)
            total.append(self.wb[sheet][f'{list(MODUL.values())[i][4]}{start}'].value)
        
        df = pd.DataFrame({
            'praktikan': praktikan,
            'nim': list_nim,
            'kelas': kelas,
            'aktif': aktif,
            'modul': mod,
            'kehadiran': kehadiran,
            'awal': awal,
            'jurnal': jurnal,
            'akhir': akhir,
            'total': total
        })
        df = df.astype({'nim': str})
        return df
    
    def get_all(self):
        res = []
        for sheet in self.kelas:
            start = 5
            while True:
                nim = self.wb[sheet][f'B{start}'].value
                if not nim:
                    break
                res.append(self.get_per_modul(sheet, nim))
                start += 1
        return pd.concat(res)
