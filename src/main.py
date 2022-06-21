import openpyxl as opxl

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

member_number = member_ws["c66"]

print(member_list)
print(fee_dict)


# ws_new = wb.create_sheet(title="")

# wb.save(r"money_data.xlsx")