import google_sheets
import random
import pyautogui
import time
import pyperclip
import argparse

from typing import Any, List, Literal

type JSON = Any

ROOM_PASSWORD = 'uwece2027'


# Generated by ChatGPT
def pad_lists_in_place(*lists: List[Any], fill_value: Any = None) -> None:
    """
    Modifies the input lists in place to make them all the same length
    by appending the specified fill value to the shorter lists.

    Args:
        *lists: Variable number of lists to modify in place.
        fill_value: Value used to pad shorter lists (default: None).
    """
    max_length = max(len(lst) for lst in lists)
    for lst in lists:
        lst += [fill_value] * (max_length - len(lst))


def spreadsheet_list_logical_or(list_a: list[str], list_b: list[str]) -> list[str]:
    # Trust me, this is necessary to avoid a bunch of extra FALSE being put below
    if len(list_a) > len(list_b):
        pad_lists_in_place(list_a, list_b)

    return [
        'TRUE' if a == 'TRUE' or b == 'TRUE' else 'FALSE'
        for a, b in zip(list_a, list_b)
    ]


class Team:
    HEADER = google_sheets.Row([
        google_sheets.Cell('Completed', bold=True),
        google_sheets.Cell('Power-up Objectives', bold=True),
        google_sheets.Cell(''),
        google_sheets.Cell('Used', bold=True),
        google_sheets.Cell('Curse', bold=True)
    ])

    def __init__(
        self,
        spreadsheet_id: str,
        power_ups: list[str],
        curses: list[str]
    ) -> None:
        self.sheet = google_sheets.GoogleSheets(spreadsheet_id)
        self.sheet_id = self.sheet.get_sheet_id_by_name('Game')
        self.power_ups = power_ups
        self.curses = curses

        self.completed_power_ups = 0
        self.used_curses = 0
        self.completed_power_ups_list = []
        self.used_curses_list = []

        sheet = google_sheets.Sheet(self.sheet_id, self.HEADER)

        for power_up in self.power_ups:
            sheet.append_row(google_sheets.Row([
                google_sheets.Cell('FALSE', checkbox=True),
                google_sheets.Cell(power_up)
            ]))
        
        self.sheet.add_sheet(sheet)
        self.sheet.clear_sheets()
        self.sheet.write()

    def update(self) -> None:
        # Check completed power-ups
        completed = self.sheet.read_list('Game', 'A')
        power_ups = self.sheet.read_list('Game', 'B')

        completed = spreadsheet_list_logical_or(completed, self.completed_power_ups_list)

        current_completed = len([cell for cell in completed if cell == 'TRUE'])
        new_curses = current_completed - self.completed_power_ups

        # Check completed curses
        curses_used = self.sheet.read_list('Game', 'D')
        curses = self.sheet.read_list('Game', 'E')

        curses_used = spreadsheet_list_logical_or(curses_used, self.used_curses_list)

        current_curses = len([cell for cell in curses_used if cell == 'TRUE'])

        if current_curses == self.used_curses and current_completed == self.completed_power_ups:
            return
    
        self.completed_power_ups = current_completed
        self.used_curses = current_curses
        self.completed_power_ups_list = completed
        self.used_curses_list = curses_used

        for i in range(new_curses):
            curses.append(random.choice(self.curses))
            curses_used.append('FALSE')

        sheet = google_sheets.Sheet(self.sheet_id, header=self.HEADER)

        # Need to extend all lists to be length of maximum
        pad_lists_in_place(completed, power_ups, curses_used, curses, fill_value='')

        for complete, power_up, curse_used, curse in zip(completed, power_ups, curses_used, curses, strict=True):
            sheet.append_row(google_sheets.Row([
                google_sheets.Cell(complete, checkbox=True if power_up != '' else False),
                google_sheets.Cell(power_up, strikethrough=complete == 'TRUE'),
                google_sheets.Cell(''),
                google_sheets.Cell(curse_used, checkbox=True if curse != '' else False),
                google_sheets.Cell(curse, strikethrough=curse_used == 'TRUE')
            ]))

        self.sheet.add_sheet(sheet)
        self.sheet.clear_sheets()
        self.sheet.write()

    sheet: google_sheets.GoogleSheets
    sheet_id: int
    power_ups: list[str]
    curses: list[str]

    completed_power_ups: int
    used_curses: int
    completed_power_ups_list: list[str]
    used_curses_list: list[str]


class Game:
    SPREADSHEET_ID = '18c59U0jZu4K_6cWjtoOvsipxxACwQNWtxkHN_I6ORZ4'

    def __init__(self, team_sheets: list[str]) -> None:
        self.sheet = google_sheets.GoogleSheets(self.SPREADSHEET_ID)
        self.objectives = self.sheet.read_list('Objectives', 'A')
        self.wild_cards = self.sheet.read_list('Wild Cards', 'A')
        self.power_ups = self.sheet.read_list('Power-ups', 'A')
        self.curses = self.sheet.read_list('Curses', 'A')
        self.board = []

        self.teams = [
            Team(team_sheet, self.power_ups, self.curses)
            for team_sheet in team_sheets
        ]

    def generate_board(self) -> None:
        # Each board contains 25 cells
        # 24 of those cells are from objectives and the last is a wild card
        # Of the 24 normal objectives, a single one is chosen as the 8-ball

        objectives = self.objectives
        while len(self.board) < 24:
            # Choose random objective
            rand = random.randint(0, len(objectives) - 1)
            self.board.append(objectives[rand])
        
        # Choose one 8-ball
        rand = random.randint(0, 23)
        self.board[rand] = f'[8-Ball] {self.board[rand]}'

        # Choose 1 wild card
        rand = random.randint(0, len(self.wild_cards) - 1)
        self.board.append(f'[WC] {self.wild_cards[rand]}')

        random.shuffle(self.board)

    def serialize_board(self) -> str:
        out = '[\n'
        for item in self.board:
            out += '\t{"name": "' + item + '"},\n'
        # Remove last comma
        out = out[:-2] + '\n'
        out += ']'
        return out
    
    def process(self) -> None:
        for team in self.teams:
            team.update() 

    sheet: google_sheets.GoogleSheets
    objectives: list[str]
    wild_cards: list[str]
    power_ups: list[str]
    curses: list[str]

    board: list[str]

    teams: list[Team]


def find_image(image: str) -> tuple[int, int]:
    while True:
        try:
            r = pyautogui.locateOnScreen(image, confidence=0.9)
            if r == None:
                raise pyautogui.ImageNotFoundException()
            else:
                return (int(r.left), int(r.top))
        except pyautogui.ImageNotFoundException:
            pass

        time.sleep(0.1)

type ClickCoordinate = Literal[
    'Room Name', 'Password', 'Nickname', 'Game', 'Board', 'Mode', 'Spectator', 'Hide', 'Make Room',
]

def get_click_coords() -> dict[ClickCoordinate, tuple[int, int]]:
    BINGOSYNC_NEW_ROOM_IMAGE = 'bingosync.png'
    x, y = find_image(BINGOSYNC_NEW_ROOM_IMAGE)

    return {
        'Room Name': (x + 250, y + 125),
        'Password': (x + 250, y + 200),
        'Nickname': (x + 250, y + 275),
        'Game': (x + 250, y + 350),
        'Board': (x + 250, y + 550),
        'Mode': (x + 250, y + 750),
        'Spectator': (x + 215, y + 930),
        'Hide': (x + 215, y + 980),
        'Make Room': (x + 550, y + 1050)
    }


def make_bingosync_room(game: Game, room_name: str):
    pyautogui.press('win')
    pyautogui.write('Google Chrome')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.write('bingosync.com')
    pyautogui.press('enter')

    coords = get_click_coords()

    pyautogui.click(coords['Room Name'])
    pyautogui.write(room_name)

    pyautogui.click(coords['Password'])
    pyautogui.write(ROOM_PASSWORD)

    pyautogui.click(coords['Nickname'])
    pyautogui.write('Host')

    pyautogui.click(coords['Game'])
    pyautogui.write('Custom')
    pyautogui.press('enter')

    pyperclip.copy(game.serialize_board())
    pyautogui.click(coords['Board'])
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(coords['Mode'])
    pyautogui.write('Lockout')
    pyautogui.press('enter')

    pyautogui.click(coords['Spectator'])
    pyautogui.click(coords['Hide'])

    pyautogui.click(coords['Make Room'])

    print('--------ROOM LOGIN INFO--------')
    print(f'Room Name: {room_name}')
    print(f'Password: {ROOM_PASSWORD}')
    print('-------------------------------')


def test():
    # TEST_SHEET_ID = 248550277
    # row = google_sheets.Row([google_sheets.Cell('FALSE', checkbox=True)])
    # sheet = google_sheets.Sheet(TEST_SHEET_ID, row, [])
    # game = Game()
    # game.sheet.add_sheet(sheet)
    # game.sheet.clear_sheets()
    # game.sheet.write()
    game = Game(['1Y_X-ubSKsUEFoBW_SuL1TUp0w2uHNQ-KScD9H9yg4w0'])

    while True:
        game.process()
        time.sleep(5)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument('--spreadsheet_ids', help='Enter team spreadsheet IDs', type=str, nargs='+')
    return parser.parse_args()


def main(spreadsheet_ids: list[str]) -> None:
    game = Game(spreadsheet_ids)
    game.generate_board()
    make_bingosync_room(game, 'WatBingo')

    # This sleep exists to avoid hitting sheets API quota, may need to increase with more teams
    time.sleep(30)

    while True:
        game.process()

        # Do not reduce this cooldown, may need to be increased with more teams
        time.sleep(5)


if __name__ == '__main__':
    try:
        main(list(parse_args().spreadsheet_ids))
    except KeyboardInterrupt:
        pass
