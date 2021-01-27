import xlsxwriter
import mysql.connector


def fetch_table_data(table_name):
    # The connect() constructor creates a connection to the MySQL server and returns a MySQLConnection object.
    cnx = mysql.connector.connect(
        host='localhost',
        database='formulir',
        user='root',
        password=''
    )

    cursor = cnx.cursor()
    sql = "SELECT * FROM datadiri "
    cursor.execute(sql)

    header = [row[0] for row in cursor.description]

    rows = cursor.fetchall()

    # Closing connection
    cnx.close()

    return header, rows


def export(datadiri):
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(datadiri + '.xlsx')
    worksheet = workbook.add_worksheet('MENU')

    # Create style for cells
    header_cell_format = workbook.add_format(
        {'bold': True, 'border': True, 'bg_color': 'yellow'})
    body_cell_format = workbook.add_format({'border': True})

    header, rows = fetch_table_data(datadiri)

    row_index = 0
    column_index = 0

    for column_name in header:
        worksheet.write(row_index, column_index,
                        column_name, header_cell_format)
        column_index += 1

    row_index += 1
    for row in rows:
        column_index = 0
        for column in row:
            worksheet.write(row_index, column_index, column, body_cell_format)
            column_index += 1
        row_index += 1

    print(str(row_index) + ' rows written successfully to ' + workbook.filename)

    # Closing workbook
    workbook.close()


# Tables to be exported
export('datadiri')
