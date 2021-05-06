from openpyxl import load_workbook


def run():
    data_book = load_workbook(filename='data.xlsx')
    db_sheet = data_book['prices']
    custom_book = load_workbook(filename='custom.xlsx')
    cb_sheet = custom_book['Лист1']

    for cb_row in cb_sheet:
        if cb_row[0].value not in ('Код', None):
            for db_row in db_sheet:
                if db_row[0].value == cb_row[0].value:
                    price = db_row[1].value
                    final_price = price * 1.04
                    cb_sheet['P' + str(cb_row[0].row)].value = round(final_price, 2)

    custom_book.save(filename='custom.xlsx')


if __name__ == '__main__':
    run()
    x = input('Done! Press "Enter".')
