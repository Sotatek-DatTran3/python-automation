from .handle_excel import HandleExcel
from .handle_query import crawl_answer


def handle_excel_command(file_path, cell_range: tuple[str, str], output_cells: list[str],  sheet_name: str = None):
    excel_handler = HandleExcel(file_path=file_path)

    print("Reading Excel file...")
    queries = excel_handler.read_excel_file(cell_range=cell_range, sheet_name=sheet_name)

    print("Crawling answers...")
    responses = crawl_answer(queries=queries)

    print("Writing back to Excel file...")
    excel_handler.write_excel_file(data=responses, output_cells=output_cells, sheet_name=sheet_name)

    print("Done.")