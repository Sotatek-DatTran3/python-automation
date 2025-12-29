from .handle_excel import HandleExcel
from .handle_query import crawl_answer
import re

def handle_excel_command(file_path, cell_range: tuple[str, str], output_cells: list[str],  sheet_name: str = None, output_sheet_name: str = None, is_headless: bool = False) -> None:
    excel_handler = HandleExcel(file_path=file_path)

    print("Reading Excel file...")
    queries = excel_handler.read_excel_file(cell_range=cell_range, sheet_name=sheet_name)

    print("Crawling answers...")
    start_cell = re.sub(r'[^0-9]', '', cell_range[0])
    responses, driver = crawl_answer(queries=queries, start=int(start_cell), is_headless=is_headless)

    print("Writing back to Excel file...")
    excel_handler.write_excel_file(data=responses, output_cells=output_cells, sheet_name=output_sheet_name)

    print("\nExcel saved. Compare data in Excel with the web, then press Enter to close the browser...")
    input("\nPress Enter to close the browser...")
    driver.quit()

    print("Done.")