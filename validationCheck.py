from openpyxl import load_workbook
from changeForm import change_form
def validationCheck(list1, list2):
    length = len(list1)
    for i in range(length):
        dict1, dict2 = list1[i], list2[i]
        if dict1 != dict2:
            print(f"Difference found at index {i}:")
            print('[result_data]')
            print(dict1)
            print('[check_target]')
            print(dict2)

excel_path = '/Users/dataly/Desktop/1E_1S_요약_평가_수정.xlsx'
save_path = '/Users/dataly/Desktop/1E_1S_요약_평가.xlsx'

change_wb = change_form(excel_path)
target_wb = load_workbook(save_path)

print(type(change_wb))
print(type(target_wb))

# if change_wb == target_wb:
#     print("validation check Done.")
# else:
#     print("validation check failed.")