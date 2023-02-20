from app import Praktikum

prak = Praktikum('Absensi dan Nilai Praktikum PBO IF 2022_2023.xlsx', 'PBO')
res = prak.get_all()

print(res.head())