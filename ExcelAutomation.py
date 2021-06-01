from openpyxl import load_workbook


DATA_FILE, DATA_SHEET_NAME = 'data.xlsx', 'prices'


def print_logo():
    logo = '''

        ██████╗ ███████╗██╗  ██╗ █████╗ 
        ██╔══██╗██╔════╝██║ ██╔╝██╔══██╗
        ██║  ██║█████╗  █████╔╝ ███████║
        ██║  ██║██╔══╝  ██╔═██╗ ██╔══██║
        ██████╔╝███████╗██║  ██╗██║  ██║
        ╚═════╝ ╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝

    '''
    print(logo)


def user_input():
    user_file = input('Enter file name (Press "Enter" for default "custom.xlsx"): ').strip()
    if user_file == '': user_file = 'custom.xlsx'
    user_sheet_name = input('Enter sheet name (Press "Enter" for default "Лист1"): ').strip()
    if user_sheet_name == '': user_sheet_name = 'Лист1'
    price_column = input('Enter price column (Press "Enter" for default "P"): ').strip().upper()
    if price_column == '': price_column = 'P'
    print_user_input(user_file, user_sheet_name, price_column)
    return user_file, user_sheet_name, price_column


def print_user_input(user_file, user_sheet_name, price_column):
    # There will be a pretty table
    print()
    print('Your file is ', user_file)
    print('Target sheet is ', user_sheet_name)
    print('Price column is ', price_column)
    print()
    print()


def insert_prices(data_file, data_sheet_name, user_file, user_sheet_name, price_column):
    data_book = load_workbook(filename=data_file)
    data_sheet = data_book[data_sheet_name]
    user_book = load_workbook(filename=user_file)
    user_sheet = user_book[user_sheet_name]

    for user_row in user_sheet:
        if user_row[0].value == None: break
        final_price = 'Didn\'t find'
        for data_row in data_sheet:
            if data_row[0].value == user_row[0].value:
                price = data_row[1].value
                final_price = round(price * 1.04, 2)
                user_sheet[price_column + str(user_row[0].row)].value = final_price
                break
        add_summary(user_row[0].value, final_price)

    user_book.save(filename=user_file)


def add_summary(code, price):
    print(code, price)


def print_summary():
    # There will be a pretty table
    print()
    print()


def run():
    print_logo()
    insert_prices(DATA_FILE, DATA_SHEET_NAME, *user_input())
    print_summary()
    input('Done! Press "Enter"...')


if __name__ == '__main__':
    run()
