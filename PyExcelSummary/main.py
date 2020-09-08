from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font


RED = 'ff4d4d'
GREEN = '66ff66'
GRAY = 'bfbfbf'
BLACK = '000001'
WHITE = 'ffffff'
ORANGE = 'ffa500'


def excelReport():
    print("Please provide the path to the file:")
    path = input()
    path = path.replace('"', '')
    wb = load_workbook(path)
    i = len(wb.sheetnames)

    failed = 0
    passed = 0
    not_performed = 0
    test_cases = 0
    passed_with_deviation = 0

    total_failed = 0
    total_passed = 0
    total_not_performed = 0
    total_test_cases = 0

    for sheet in wb.worksheets:
        wb = load_workbook(path)
        currentSheetName = wb.sheetnames[len(wb.sheetnames) - i]
        ws = wb[wb.sheetnames[len(wb.sheetnames) - i]]
        if currentSheetName == "DocHistory" or currentSheetName == "Test_Summary" or currentSheetName == "Status":
            i += 1
            continue
        else:
            for rows in ws.iter_rows():
                for cell in rows:
                    if str(cell.value).lower() == 'failed':
                        cell.fill = PatternFill(start_color=RED, end_color=RED, fill_type="solid")
                        cell.font = Font(bold=True, color=GREEN, size=16)
                        total_failed += 1
                        total_test_cases += 1
                        test_cases += 1
                        failed += 1
                    if str(cell.value).lower() == 'passed':
                        if str(ws.cell(row=cell.row, column=cell.column + 1).value).lower() != 'none':
                            passed_with_deviation += 1
                        cell.fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
                        cell.font = Font(bold=True, color=RED, size=16)
                        total_passed += 1
                        total_test_cases += 1
                        test_cases += 1
                        passed += 1
                    if str(cell.value).lower() == 'not performed':
                        cell.fill = PatternFill(start_color=GRAY, end_color=GRAY, fill_type="solid")
                        cell.font = Font(bold=True, color=WHITE, size=16)
                        total_not_performed += 1
                        total_test_cases += 1
                        test_cases += 1
                        not_performed += 1
        print(f"{currentSheetName} \n 'passed:{passed} (minor deviation: {passed_with_deviation})' 'failed:{failed}' 'not_performed:{not_performed}' \n")

        ws = wb['Test_Summary']
        for rows in ws.iter_rows():
            for cell in rows:
                currentSheetName = wb.sheetnames[len(wb.sheetnames) - i]
                if str(cell.value) == str(currentSheetName):
                    ws.cell(row=cell.row, column=cell.column + 1).value = passed
                    ws.cell(row=cell.row, column=cell.column + 2).value = passed_with_deviation
                    ws.cell(row=cell.row, column=cell.column + 3).value = failed
                    ws.cell(row=cell.row, column=cell.column + 4).value = not_performed
                    ws.cell(row=cell.row, column=cell.column + 5).value = test_cases
                if str(cell.value) == 'Overall test result':
                    if total_failed > 0:
                        ws.cell(row=cell.row, column=cell.column+1).fill = PatternFill(start_color=RED, end_color=RED, fill_type="solid")
                        ws.cell(row=cell.row, column=cell.column + 1).value = "Failed"
                    elif passed_with_deviation > 0:
                        ws.cell(row=cell.row, column=cell.column + 1).fill = PatternFill(start_color=ORANGE, end_color=ORANGE,
                                                                                         fill_type="solid")
                        ws.cell(row=cell.row, column=cell.column + 1).value = "Passed with minor deviation."
                    else:
                        ws.cell(row=cell.row, column=cell.column + 1).fill = PatternFill(start_color=GREEN,
                                                                                         end_color=GREEN,
                                                                                         fill_type="solid")
                        ws.cell(row=cell.row, column=cell.column + 1).value = "Passed."
                        
        i += 1

        wb.save(path)
        passed_with_deviation = 0
        failed = 0
        passed = 0
        not_performed = 0
        test_cases = 0

    print(f"Done.")

if __name__ == '__main__':
    main()
