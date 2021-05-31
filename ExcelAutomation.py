from openpyxl import load_workbook


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
    dataFile, dataSheet = 'data.xlsx', 'prices'
    userFile = input('Enter file name (Press "Enter" for "custom.xlsx"): ').strip()
    if userFile == '': userFile = 'custom.xlsx'
    userSheet = input('Enter sheet name (Press "Enter" for "Лист1"): ').strip()
    if userSheet == '': userSheet = 'Лист1'
    price_column = 'P'
    return dataFile, dataSheet, userFile, userSheet, price_column


def insert_prices(dataFile, dataSheet, userFile, userSheet, price_column):
    data_book = load_workbook(filename=dataFile)
    db_sheet = data_book[dataSheet]
    user_book = load_workbook(filename=userFile)
    ub_sheet = user_book[userSheet]

    for ub_row in ub_sheet:
        if ub_row[0].value not in ('Код', None):
            for db_row in db_sheet:
                if db_row[0].value == ub_row[0].value:
                    price = db_row[1].value
                    final_price = price * 1.04
                    ub_sheet[price_column + str(ub_row[0].row)].value = round(final_price, 2)

    user_book.save(filename=userFile)


def run():
    print_logo()
    insert_prices(*user_input())
    input('Done! Press "Enter".')


if __name__ == '__main__':
    run()

