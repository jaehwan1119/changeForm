import time
from changeForm import change_form

start_time = time.time()

# excel 파일 path
excel_path = '/Users/dataly/Desktop/1E_1S_요약_평가_수정.xlsx'
save_path = '/Users/dataly/Desktop/1E_1S_요약_평가.xlsx'

change_form(excel_path).save(save_path)

end_time = time.time()
print(end_time - start_time)