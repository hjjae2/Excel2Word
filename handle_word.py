from mailmerge import MailMerge

import openpyxl

def create_sender_table(file_name, sender, receiver, r_date):
    template_file = "./demo.docx"
    file_path = file_name
    target_name = file_path.split("/")[-1].split(".")[0]
    target_name = target_name + "_거래내역서.docx"
    target_path = file_path.split("/")[:-1]
    target_path = "/".join(target_path)
    target_path = target_path + "/" + target_name
    print("File_path:", file_path)
    print("Target_path:", target_path)

    excel_document = openpyxl.load_workbook(file_path)
    sheet_name = excel_document.get_sheet_names()[0]
    sheet = excel_document.get_sheet_by_name(sheet_name)

    print("Active Sheet:", sheet_name)

    # read excel file.
    sales_history = []
    total_price = 0
    last_idx = 0;
    flag = 1
    all_rows = sheet.rows
    for idx, row in enumerate(all_rows, start=-1):
        # loop handling 1
        if idx > 1000:
            flag = -1
            break

        last_idx = idx

        if idx < 1:
            continue

        sales_item = {}
        sales_item["no"] = str(idx)
        sales_item["date"] = str(row[0].value)
        sales_item["date"] = "/".join(sales_item["date"].split("-")[1:3])
        sales_item["date"] = sales_item["date"].split(" ")[0]
        # loop handling 2
        print(sales_item["date"], end=" ")
        if sales_item["date"] == "None" or sales_item["date"] == "":
            flag = 0
            break
        sales_item["work"] = row[3].value
        sales_item["content"] = ""
        # price column exception handling.
        if row[4].value is None:
            total_price += 0
            sales_item["price"] = "-"
        elif type(row[4].value) is str:
            total_price += 0
        else:
            sales_item["price"] = format(row[4].value, ",")
            print(sales_item["price"])
            total_price += row[4].value
        sales_item["tax"] = row[7].value
        sales_history.append(sales_item)

    print("Reading Data >>>")
    if flag == 1:
        print("\tSUCCESS: reading excel data")
    elif flag == 0:
        print("\tWARNING: for-loop's idx reaches until date is None.")
    else:
        print("\tERROR: for-loop's idx > 1000.")

    # write word file.
    total_price = format(total_price, ",")

    print("Sum >>>")
    print("\tTotal_Data:", last_idx, "건")
    print("\tTotal_Price:", total_price, "원")

    print("Write FIle >>>")
    flag = 1
    try:
        document = MailMerge(template_file)
        document.merge(
            number=sender[0],
            company=sender[1],
            name=sender[2],
            address=sender[3],
            call=sender[4],
            fax=sender[5],
            business=sender[6],
            category=sender[7]
        )
        document.merge(
            receiver=receiver,
            r_date=r_date,
            total_price=total_price
        )
        document.merge_rows('no', sales_history)
        document.write(target_path)
    except Exception as e:
        flag = 0
        print("\tError:", e)
    if flag == 1:
        print("\tSUCCESS:Writing file ->", target_path)