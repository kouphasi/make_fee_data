from hashlib import new
import openpyxl as opxl
from unicodedata import name
import openpyxl as opxl
import openpyxl
import math
from action.indiv_fee import indiv_fee
import datetime

wb = opxl.load_workbook(r"money_data.xlsx",data_only=True)

fee_ws = wb["内訳"]

member_ws = wb["名簿"]

fee_dict = {}
for i in range(9):
    fee_dict[str(fee_ws[f"a{i+1}"].value)] = fee_ws[f"b{i+1}"].value

member_list = []
for i in range(2,53):
    member_list.append(
        {
            member_ws["a1"].value:member_ws[f"a{i}"].value,
            member_ws["b1"].value:member_ws[f"b{i}"].value,
            member_ws["c1"].value:member_ws[f"c{i}"].value,
            member_ws["d1"].value:member_ws[f"d{i}"].value,
            member_ws["e1"].value:member_ws[f"e{i}"].value,
            member_ws["f1"].value:member_ws[f"f{i}"].value,
            member_ws["g1"].value:member_ws[f"g{i}"].value,
            member_ws["h1"].value:member_ws[f"h{i}"].value,
            member_ws["i1"].value:member_ws[f"i{i}"].value,
            member_ws["j1"].value:member_ws[f"j{i}"].value,
            member_ws["k1"].value:member_ws[f"k{i}"].value,
            member_ws["l1"].value:member_ws[f"l{i}"].value
        }
    )

member_number = member_ws["c66"].value

# print(member_list)
# print(fee_dict)

new_wb = opxl.Workbook()
payment_ws = new_wb["Sheet"]
payment_ws.title="支払い名簿"

for member in member_list:
    member["料金徴収"] = indiv_fee(new_wb,fee_dict,member,member_number)


payment_ws["a1"] = "学年"
payment_ws["b1"] = "名前"
payment_ws["c1"] = "金額"

for index,member in enumerate(member_list,2):
    grade = member["学年"]
    payment_ws[f"a{index}"] = grade
    payment_ws[f"b{index}"] = member["名前"]
    payment_ws[f"c{index}"] = member["料金徴収"]
    
new_wb.save(f"{datetime.date.today()}-明細表.xlsx")

wb.save(r"money_data.xlsx")