from datetime import datetime as dt
from plistlib import InvalidFileException  # Import datetime class and rename it to avoid conflicts
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

class ExcelService:
    path_to_file = "data.xlsx"


    # @staticmethod
    # def create_file():
    #     try:
    #         # Try to load an existing workbook
    #         ExcelService.wb = load_workbook(ExcelService.path_to_file)
    #     except FileNotFoundError:
    #         # If the file doesn't exist, create a new workbook
    #         ExcelService.wb = Workbook()


    @staticmethod
    def create_file():
        try:
            # Try to open an existing workbook
            ExcelService.wb = load_workbook(ExcelService.path_to_file)
        except (FileNotFoundError, InvalidFileException):
            # If the file doesn't exist or is invalid, create a new workbook
            ExcelService.wb = Workbook()



    '''
    @staticmethod
    def create_tab(tab_name):
        try:
            # Try to get an existing sheet
            ExcelService.sheet = ExcelService.wb[tab_name]
        except KeyError:
            # If the sheet doesn't exist, create a new one
            ExcelService.sheet = ExcelService.wb.create_sheet(title=tab_name)

            # Write headers to the new sheet
            for col_num, header in enumerate(ExcelService.headers, 1):
                ExcelService.sheet.cell(row=1, column=col_num, value=header)

    '''

    @staticmethod
    def create_tab(tab_name):
        try:
            # Try to get an existing sheet
            ExcelService.sheet = ExcelService.wb[tab_name]
        except KeyError:
            # If the sheet doesn't exist, create a new one
            ExcelService.sheet = ExcelService.wb.create_sheet(title=tab_name)

            # Write headers to the new sheet only if it's a new sheet
            if ExcelService.sheet.max_row == 1:
                for col_num, header in enumerate(ExcelService.headers, 1):
                    ExcelService.sheet.cell(row=1, column=col_num, value=header)



    @staticmethod
    def write_to_file(tab_name, aiq_item):
        # Create a new sheet or get the existing one
        ExcelService.create_tab(tab_name)

        # Get the current date in the specified format
        current_date = dt.now().strftime("%d-%m-%y %H:%M")

        # Get the last row number
        last_row = ExcelService.sheet.max_row + 1

        # Write data to the sheet
        ExcelService.sheet.cell(row=last_row, column=1, value=current_date)
        ExcelService.sheet.cell(row=last_row, column=2, value=aiq_item.date)
        ExcelService.sheet.cell(row=last_row, column=3, value=aiq_item.pm2_5)
        ExcelService.sheet.cell(row=last_row, column=4, value=aiq_item.pm10)
        ExcelService.sheet.cell(row=last_row, column=5, value=aiq_item.temperature)
        ExcelService.sheet.cell(row=last_row, column=6, value=aiq_item.humidity)
        ExcelService.sheet.cell(row=last_row, column=7, value=aiq_item.pressure)
        ExcelService.sheet.cell(row=last_row, column=8, value=aiq_item.heca_temperature)
        ExcelService.sheet.cell(row=last_row, column=9, value=aiq_item.heca_humidity)

        # Save the workbook to a file
        ExcelService.wb.save(ExcelService.path_to_file)

        print(f"Processed {tab_name} location")

    @staticmethod
    def format_tab(tab_name, font_size=12, bold=True, align_center=True):       

        # Apply formatting options
        for col_num in range(1, len(ExcelService.headers) + 1):
            cell = ExcelService.sheet.cell(row=1, column=col_num)
            cell.font = Font(size=font_size, bold=bold)
            ExcelService.sheet.column_dimensions[get_column_letter(col_num)].width = 25
            if align_center:
                cell.alignment = Alignment(horizontal='center')

    # Headers for the Excel sheets
    headers = ["Updated Timestamp", "Fetched Timestamp", "PM 2,5", "PM 10", "Temperature", "Humidity", "Pressure", "HECA Temperature", "HECA Humidity"]

    
