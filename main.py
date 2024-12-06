from openpyxl import load_workbook, Workbook
from datetime import datetime


class AdmissionDiffChecker:
    def __init__(self):
        self.old_admission_map = set()
        self.header_row = None
        self.diff_rows = []

    def load_older_worksheet(self):
        print("Parsing Older Worksheet")
        wb = load_workbook(filename="assets/example_sheet_A.xlsx")
        ws = wb.active
        for row in ws.iter_rows(min_row=1, values_only=True):
            # Write header row for result
            if not self.header_row:
                self.header_row = row
                continue
            if not row[0] or not row[1]:
                continue
            self.old_admission_map.add(row[0] + row[1])

    def load_and_compare_newer_worksheet(self):
        print("Parsing Newer Worksheet")
        wb = load_workbook(filename="assets/example_sheet_B.xlsx")
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0] or not row[1]:
                continue
            name = row[0] + row[1]
            if name not in self.old_admission_map:
                self.diff_rows.append(row)
            # This is for de-duplication
            self.old_admission_map.add(name)

    def write_result_to_new_worksheet(self):
        wb = Workbook()
        ws = wb.active
        ws.append(checker.header_row)
        for row in self.diff_rows:
            ws.append(row)
        wb.save("assets/admin_diff.xlsx")

    def do_comparison(self):
        self.load_older_worksheet()
        self.load_and_compare_newer_worksheet()
        self.write_result_to_new_worksheet()


start_time = datetime.now()
print(f"Begun Checking Diffs {start_time}")
checker = AdmissionDiffChecker()
checker.do_comparison()
finish_time = datetime.now() - start_time
print(f"Finished Checking Diffs. Time Taken: {finish_time.microseconds} microseconds")
