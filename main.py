import json
import time
import pandas as pd
from changeForm import change_form
# from validationCheck import validationCheck

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

start_time = time.time()

# excel 파일 path
excel_path = '/Users/dataly/Desktop/1E_1S_요약_평가_수정.xlsx'
# save_path = '/Users/dataly/Desktop/1E_1S_요약_평가.xlsx'

change_form(excel_path)

end_time = time.time()
print(end_time - start_time)