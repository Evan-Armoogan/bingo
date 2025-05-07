import json
import typing
import string
from datetime import datetime
from pathlib import Path

from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import HttpError

PACKAGE_DIR = Path(__file__).parent

if typing.TYPE_CHECKING:
    # As per sheets API docs, pylance doesn't like this but it is required
    from googleapiclient._apis.sheets.v4 import SheetsResource  # type: ignore

SERV_ACC_PATH = PACKAGE_DIR / 'GServAcc'  # NOTE: Modify as necessary

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'

class RGBA:
    def __init__(self, red: int, green: int, blue: int, alpha: int):
        self.red = red
        self.green = green
        self.blue = blue
        self.alpha = alpha

    red: int
    green: int
    blue: int
    alpha: int

# Use these websites when adding colours:
#   https://redketchup.io/color-picker
#   https://rgbacolorpicker.com/hex-to-rgba
type Colour = typing.Literal['White', 'Red', 'Blue']
COLOUR_DICT: dict[Colour, RGBA] = {
    'White': RGBA(255, 255, 255, 1),
    'Black': RGBA(0, 0, 0, 1),
    'Red': RGBA(234, 153, 153, 1),
    'Blue': RGBA(164, 194, 244, 1)
}

type CellAlignment = typing.Literal['Left', 'Center', 'Right']
ALIGNMENT_DICT: dict[CellAlignment, str] = {
    'Left': 'LEFT',
    'Center': 'CENTER',
    'Right': 'RIGHT'
}

type JSON = typing.Any


class Cell:
    @staticmethod
    def format_data_str(data: str) -> str:
        # Refer to this string as an example of how we must format to make the CSV and JSON happy
        # f'=HYPERLINK("{TBA_LINK_PREFIX}/event/{division}#rankings", "Full Rankings")' must become
        # f'\\"=HYPERLINK(\\"\\"{TBA_LINK_PREFIX}/event/{division}#rankings\\"\\", \\"\\"Full Rankings\\"\\")\\"'
        # We reformat strings in this manner in this function

        comma = data.count(',') > 0 or data.count('-') > 0
        escape = '\\"\\' if comma else '\\'

        offset = 0
        while (x := data.find('"', offset)) != -1:
            data = data[:x] + escape + data[x:]
            offset = x + len(escape) + 1  # +1 so we move past this comma to next char

        if comma:
            data = f'\\"{data}\\"'
        
        return data

    def __init__(
        self,
        data: str,
        sheet_id: int = 0,
        row: int = -1,
        column: int = -1,
        font_size: int = 10,
        text_colour: Colour = 'Black',
        cell_colour: Colour = 'White',
        bold: bool = False,
        italic: bool = False,
        length: int = 1,
        alignment: CellAlignment = 'Left',
        datetime: str | None = None,
        checkbox: bool = False,
        strikethrough: bool = False
    ) -> None:
        if length < 1:
            raise NotImplementedError('Cell length must be at least one')

        self.data = Cell.format_data_str(data)
        self.sheet_id = sheet_id
        self.row = row
        self.column = column
        self.font_size = font_size
        self.text_colour = text_colour
        self.cell_colour = cell_colour
        self.bold = bold
        self.italic = italic
        self.length = length
        self.alignment = alignment
        self.datetime = datetime
        self.checkbox = checkbox
        self.strikethrough = strikethrough

    def __eq__(self, other):
        if not isinstance(other, Cell):
            raise NotImplementedError()
        else:
            return (
                self.data == other.data and
                self.sheet_id == other.sheet_id and
                self.row == other.row and
                self.column == other.column and
                self.font_size == other.font_size and
                self.text_colour == other.text_colour and
                self.cell_colour == other.cell_colour and
                self.bold == other.bold and
                self.italic == other.italic and
                self.length == other.length and
                self.alignment == other.alignment and
                self.datetime == other.datetime and
                self.checkbox == other.checkbox and
                self.strikethrough == other.strikethrough
            )

    def get_format_request(self) -> JSON:
        with open(PACKAGE_DIR / 'format_request_template.json', 'r', encoding='utf-8') as f:
            template = string.Template(f.read())
        return json.loads(template.substitute(
            SheetId = self.sheet_id,
            StartRowIndex = self.row,
            EndRowIndex = self.row + 1,
            StartColIndex = self.column,
            EndColIndex = self.column + 1,
            BackgroundRed = str(float(COLOUR_DICT.get(self.cell_colour).red/255)),
            BackgroundGreen = str(float(COLOUR_DICT.get(self.cell_colour).green/255)),
            BackgroundBlue = str(float(COLOUR_DICT.get(self.cell_colour).blue/255)),
            BackgroundAlpha = str(float(COLOUR_DICT.get(self.cell_colour).alpha)),
            TextAlignment = ALIGNMENT_DICT.get(self.alignment),
            TextRed = str(float(COLOUR_DICT.get(self.text_colour).red/255)),
            TextGreen = str(float(COLOUR_DICT.get(self.text_colour).green/255)),
            TextBlue = str(float(COLOUR_DICT.get(self.text_colour).blue/255)),
            TextAlpha = str(float(COLOUR_DICT.get(self.text_colour).alpha)),
            FontSize = self.font_size,
            TextBold = 'true' if self.bold else 'false',
            TextItalic = 'true' if self.italic else 'false',
            Strikethrough = 'true' if self.strikethrough else 'false'
        ))
    
    def get_checkbox_request(self) -> JSON:
        with open(PACKAGE_DIR / 'checkbox_request_template.json', 'r', encoding='utf-8') as f:
            template = string.Template(f.read())
        return json.loads(template.substitute(
            SheetId = self.sheet_id,
            StartRowIndex = self.row,
            EndRowIndex = self.row + 1,
            StartColIndex = self.column,
            EndColIndex = self.column + 1
        ))
    
    def get_merge_request(self) -> JSON:
        with open(PACKAGE_DIR / 'merge_cell_request_template.json', 'r', encoding='utf-8') as f:
            template = string.Template(f.read())
        return json.loads(template.substitute(
            SheetId = self.sheet_id,
            StartRowIndex = self.row,
            EndRowIndex = self.row + 1,
            StartColIndex = self.column,
            EndColIndex = self.column + self.length
        ))
    
    def get_datetime_request(self) -> JSON:
        if self.datetime is not None:
            with open(PACKAGE_DIR / 'datetime_format_request_template.json', 'r', encoding='utf-8') as f:
                template = string.Template(f.read())
            return json.loads(template.substitute(
                SheetId = self.sheet_id,
                StartRowIndex = self.row,
                EndRowIndex = self.row + 1,
                StartColIndex = self.column,
                EndColIndex = self.column + 1,
                DateTimePattern = self.datetime
            ))
        else:
            return {}
    
    def get_requests(self) -> list[JSON]:
        ret = []
        DEFAULT_FORMAT = Cell(data=self.data, sheet_id=self.sheet_id, row=self.row, column=self.column)
        if self == DEFAULT_FORMAT:
            return ret
        
        if self.length != 1:
            # NOTE: Length < 1 prevented in constructor
            # Length > 1 means we need to merge cells
            ret.append(self.get_merge_request())
        
        if self.datetime != None:
            # We need to format the cell as date/time with the given pattern
            ret.append(self.get_datetime_request())

        if self.checkbox:
            ret.append(self.get_checkbox_request())

        # Format is different
        ret.append(self.get_format_request())

        return ret


    sheet_id: int
    row: int
    column: int

    data: str
    font_size: int
    text_colour: Colour
    cell_colour: Colour
    bold: bool
    italic: bool
    length: int
    alignment: CellAlignment
    datetime: str | None
    checkbox: bool
    strikethrough: bool


class Row:
    def __init__(self, sheet_id: int = 0, row: int = -1) -> None:
        self.cells = []
        self.sheet_id = sheet_id
        self.row = row
    
    def __init__(self, cells: list[Cell], sheet_id: int = 0, row: int = -1) -> None:
        self.sheet_id = sheet_id
        self.row = row
        self.cells = []
        for cell in cells:
            self.append_cell(cell)

    def get_csv_data(self):
        csv = ''
        for cell in self.cells:
            csv += cell.data + ','

        # Remove trailing comma if data in csv
        return csv[:-1] if csv != '' else ','
    
    def get_data_request(self) -> JSON:
        with open(PACKAGE_DIR / 'data_request_template.json', 'r', encoding='utf-8') as f:
            template = string.Template(f.read())
        return json.loads(template.substitute(
            CsvData = self.get_csv_data(),
            SheetId = self.sheet_id,
            RowIndex = self.row
        ))

    def get_requests(self) -> list[JSON]:
        ret = [self.get_data_request()]
        for cell in self.cells:
            ret += cell.get_requests()
        return ret
    
    def insert_cell(self, cell: Cell, col: int) -> None:
        cell.sheet_id = self.sheet_id
        cell.row = self.row
        cell.column = col

        for cell in self.cells[col:]:
            cell.column += cell.length

        self.cells.insert(col, cell)

        # We need to add blank cells that will be merged with the one we added when we call sheets API
        for i in range(1, cell.length):
            self.cells.insert(col + i, Cell('', sheet_id=self.sheet_id, row=self.row, column=col+i))

    def append_cell(self, cell: Cell) -> None:
        self.insert_cell(cell, len(self.cells))

    def prepend_cell(self, cell: Cell) -> None:
        self.insert_cell(cell, 0)

    def set_row(self, row: int) -> None:
        for cell in self.cells:
            cell.row = row
        self.row = row
    
    def set_sheet_id(self, sheet_id: int) -> None:
        for cell in self.cells:
            cell.sheet_id = sheet_id
        self.sheet_id = sheet_id

    sheet_id: int
    row: int
    cells: list[Cell]


class Sheet:
    def __init__(self, sheet_id: int, header: Row, rows: list[Row] = []) -> None:
        self.sheet_id = sheet_id
        self.header = header
        
        self.header.set_sheet_id(sheet_id)
        self.header.set_row(0)
        self.rows = []
        for row in rows:
            self.append_row(row)

    def get_freeze_request(self) -> JSON:
        with open(PACKAGE_DIR / 'freeze_row_template.json', 'r', encoding='utf-8') as f:
            template = string.Template(f.read())
        return json.loads(template.substitute(
            SheetId = self.sheet_id,
            NumFrozenRows = 1
        ))

    def get_requests(self) -> list[JSON]:
        ret = [self.get_freeze_request()] + self.header.get_requests()
        for row in self.rows:
            ret += row.get_requests()
        return ret
    
    def insert_row(self, row: Row, row_num: int) -> None:

        row_data = Row(row.cells, self.sheet_id, row_num + 1)  # +1 for header

        for row in self.rows[row_num:]:
            row.row += 1
            for cell in row.cells:
                cell.row += 1
        
        self.rows.insert(row_num, row_data)
    
    def append_row(self, row: Row):
        self.insert_row(row, len(self.rows))
    
    def prepend_row(self, row: Row):
        self.insert_row(row, 0)

    sheet_id: int
    header = Row
    rows: list[Row]


class GoogleSheets:
    @staticmethod
    def __activate_service() -> 'SheetsResource':
        with open(SERV_ACC_PATH, 'r', encoding='utf-8') as f:
            GServAcc = json.loads(f.read())

        creds = ServiceAccountCredentials.from_json_keyfile_dict(GServAcc, scopes=SCOPES)
        return build('sheets', 'v4', credentials=creds)

    def __init__(self, spreadsheet_id: str) -> None:
        self.service = GoogleSheets.__activate_service()
        self.spreadsheet_id = spreadsheet_id
        self.sheets = []

    def get_requests(self) -> list[str]:
        ret = []
        for sheet in self.sheets:
            ret += sheet.get_requests()
        return ret
    
    def get_request_body(self) -> dict[str, typing.Any]:
        body = {
            "requests": self.get_requests(),
            "includeSpreadsheetInResponse": False
        }
        return body
    
    # NOTE: The following function was generated by ChatGPT
    def clear_sheets(self) -> None:
        """
        Resets all sheets in the spreadsheet:
        - Clears all data
        - Clears all formatting
        - Unmerges merged cells
        - Unfreezes any frozen rows or columns
        """

        sheets_to_find = [sheet.sheet_id for sheet in self.sheets]

        # Step 1: Get all sheets
        try:
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
        except HttpError as e:
            print(e)
            return
        sheets = spreadsheet.get('sheets', [])
        sheets = [sheet for sheet in sheets if sheet['properties']['sheetId'] in sheets_to_find]

        requests = []

        for sheet in sheets:
            sheet_id = sheet['properties']['sheetId']
            sheet_name = sheet['properties']['title']

            # a) Unmerge all cells
            requests.append({
                'unmergeCells': {
                    'range': {
                        'sheetId': sheet_id
                    }
                }
            })

            # b) Clear formatting
            requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': sheet_id
                    },
                    'cell': {
                        'userEnteredFormat': {}
                    },
                    'fields': 'userEnteredFormat'
                }
            })

            # c) Unfreeze rows/columns
            requests.append({
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': sheet_id,
                        'gridProperties': {
                            'frozenRowCount': 0,
                            'frozenColumnCount': 0
                        }
                    },
                    'fields': 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount'
                }
            })

            # d) Clear cell values
            try:
                self.service.spreadsheets().values().clear(
                    spreadsheetId=self.spreadsheet_id,
                    range=sheet_name
                ).execute()
            except HttpError as e:
                print(e)
                return

        # Step 2: Batch update formatting changes
        if requests:
            try:
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={'requests': requests}
                ).execute()
            except HttpError as e:
                print(e)
                return
    
    def write(self) -> None:
        self.clear_sheets()
        # json.dump(self.get_request_body(), open('query.json', 'w'), sort_keys=True, indent='\t', separators=(',', ': '))  # For testing
        try:
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id, body=self.get_request_body()
            ).execute()
        except HttpError as e:
            print(e)
            return
    
    def add_sheet(self, sheet: Sheet) -> None:
        self.sheets.append(sheet)

    def read_list(self, sheet: str, column: str) -> list[str]:
        range_name = f'{sheet}!{column}2:{column}'

        try:
            result = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id,
                                                            range=range_name).execute()
        except HttpError as e:
            print(e)
            return

        rows: list[list[str]] = result.get('values', [])
        return [val[0] for val in rows]

    # Generated by ChatGPT, minor adaptations made by me
    def get_sheet_id_by_name(self, sheet_name: str) -> int:
        """
        Returns the sheet ID for the sheet with the given name.
        
        Args:
            sheet_name (str): The name of the sheet to look for.

        Returns:
            int: The sheet ID if found, or None if not found.
        """
        try:
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
        except HttpError as e:
            print(e)
            return

        for sheet in spreadsheet.get('sheets', []):
            properties = sheet.get('properties', {})
            if properties.get('title') == sheet_name:
                return properties.get('sheetId')
        
        raise RuntimeError(f'Unable to find sheet "{sheet_name}" in spreadsheet')

    service: 'SheetsResource'
    spreadsheet_id: str
    sheets: list[Sheet]



def test() -> None:
    header = Row([
        Cell('Match'),
        Cell('Red', cell_colour='Red', length=3, alignment='Center'),
        Cell('Red Score', cell_colour='Red'),
        Cell('Blue Score', cell_colour='Blue'),
        Cell('Blue', cell_colour='Blue', length=3, alignment='Center'),
        Cell('TBA Breakdown'),
        Cell('Time (ET)')
    ])

    sheet = Sheet(72756875, header)
    sheet.append_row([
            Cell('Newton Qualification 1'),
            Cell('2056', cell_colour='Red', bold=True, italic=True),
            Cell('1323', cell_colour='Red', italic=True),
            Cell('1678', cell_colour='Red', italic=True),
            Cell('300', cell_colour='Red', bold=True),
            Cell('251', cell_colour='Blue'),
            Cell('118', cell_colour='Blue', italic=True),
            Cell('1114', cell_colour='Blue', bold=True, italic=True),
            Cell('254', cell_colour='Blue', italic=True),
            Cell('https://www.thebluealliance.com/match/2025oncmp_f1m1'),
            Cell(datetime.fromtimestamp(1744068219).isoformat(), datetime='ddd h:mm AM/PM')
    ])

    SPREADSHEET_ID = '1V6LoRB9yLSjy9yhGqRE22ge9v2v3FHlIk6pwzfX_1-Q'
    spreadsheet = GoogleSheets(SPREADSHEET_ID)
    spreadsheet.add_sheet(sheet)
    spreadsheet.write()


if __name__ == "__main__":
    test()
