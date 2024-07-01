import time
from changeForm import change_form

start_time = time.time()

fileList = ['A', 'B', 'C']

for name in fileList:
    # excel 파일 path
    excel_path = f'/Users/dataly/Desktop/24년 요약 평가 작업 1E_1S (24.06.03~06.10) ({name}팀)_198건.xlsx'
    save_path = f'/Users/dataly/Desktop/24년 요약 평가 작업 1E_1S (24.06.03~06.10) ({name}팀)_198건_result.xlsx'
    sheet_base = f'SUMVAL_{name}_'

    change_form(excel_path, sheet_base).save(save_path)

end_time = time.time()
print(end_time - start_time)