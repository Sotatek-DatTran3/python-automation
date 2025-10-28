from openpyxl import load_workbook
import re

class HandleExcel():
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.workbook = None

    def __split_cell_address(self, cell):
        match = re.match(r"([A-Z]+)([0-9]+)", cell)
        if not match:
            raise ValueError("Invalid cell format")
        column, row = match.groups()
        return column, int(row)

    def __is_same_column(self, cell1, cell2):
        col_from = re.match(r"([A-Z]+)", cell1)
        col_to = re.match(r"([A-Z]+)", cell2)
        
        if not col_from or not col_to:
            raise ValueError("Invalid cell format")

        if col_from.group(0) != col_to.group(0):
            raise ValueError("Cells must be in the same column")

    def read_excel_file(self, cell_range: tuple[str, str], sheet_name: str = None):
        try:
            from_cell = cell_range[0]
            to_cell = cell_range[1]
            self.__is_same_column(from_cell, to_cell)

            self.workbook = load_workbook(filename=self.file_path, data_only=True)
            sheet = self.workbook.active if sheet_name is None else self.workbook[sheet_name]

            data = []
            for row in sheet[f"{from_cell}:{to_cell}"]:
                for cell in row:
                    if cell.value is not None:
                        text = cell.value.replace("質問：\n", "", 1).strip()
                        data.append(text)
            
            return data
        except Exception as e:
            print(f"Error reading the Excel file: {e}")
    
    def write_excel_file(self, data,  output_cells: list[str], sheet_name: str = None):
        column1, start_row1 = self.__split_cell_address(output_cells[0])
        column2, start_row2 = self.__split_cell_address(output_cells[1])
        sheet = self.workbook.active if sheet_name is None else self.workbook[sheet_name]

        for idx in range(len(data)):
            cell1 = f"{column1}{start_row1 + idx}"
            cell2 = f"{column2}{start_row2 + idx}"
            sheet[cell1] = data[idx]["answer"]
            text = ""
            for idx, file in enumerate(data[idx]["references"]):
                text += f"{idx + 1}. {file}\n"
            sheet[cell2] = text.strip()

        self.workbook.save(self.file_path)