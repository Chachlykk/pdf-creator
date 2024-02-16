import win32com.client as win32client
import read_settings


def main():
    print("execution...\n")

    settings = read_settings.read_settings()

    excel = win32client.Dispatch("Excel.Application")

    # open sheets in workbook
    workbook = excel.Workbooks.Open(settings["file_name"])
    worksheet = workbook.Sheets(settings["template_page"])
    worksheet_mail = workbook.Sheets(settings["pdf_name_page"])

    # reading settings
    results_template = settings["results_cells_template"]
    members = settings["member_page"]
    results = settings["results_columns"]
    file_name_column = settings["column_with_file_name"]
    first_member = settings["first_member_number"]
    last_member = settings["last_member_number"]

    while first_member <= last_member:
        # filling out the template
        for i in range(len(results_template)):
            worksheet.Range(results_template[i]).Formula = "='" + members + "'!" + results[i] + str(first_member)

        worksheet.Range(settings["code_cell_template"]).Formula = \
            ("='" + members + "'!" + settings["code_column"] + str(first_member))
        worksheet.Range(settings["name_cell_template"]).Formula =\
            "='" + members + "'!" + settings["name_column"] + str(first_member)

        # creating pdf file from the template
        worksheet.ExportAsFixedFormat(0, settings["folder_with_pdf"] + worksheet_mail.Range(
            file_name_column + str(first_member)).Value)
        first_member += 1
    workbook.Close(SaveChanges=False)
    excel.Quit()

    input("Successfully completed!\n")


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(e)
        input()
