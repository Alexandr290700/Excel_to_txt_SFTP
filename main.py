import os
import openpyxl
import pysftp


sftp_host = 'hostname'
sftp_port = 22
sftp_username = 'username'
sftp_password = 'password'
sftp_directory = '/path/to/directory/'


# Путь к директ. к excel файлам
excel_directory = '/path/to/excel/files/'


excel_files = [f for f in os.listdir(excel_directory) if f.endswith('.xlsx')]

# Перебор Excel файлов
for excel_file in excel_files:
    # Открываем файл
    wb = openpyxl.load_workbook(os.path.join(excel_directory, excel_file))
    sheet = wb.active

    # Создание текстового файла
    output_file = open('output.txt', 'w')


    # Обходим строки в excel и записываем в текст. файл
    for row in sheet.iter_rows(min_row=2, values_only=True):
        invoice_type = row[0]
        if invoice_type == "H":
            data = "\t".join(str(val) for val in row[1:])
            line = f"{invoice_type}\t{data}\n"
        elif invoice_type == "I":
            data = "\t".join(str(val) for val in row[1:])
            line = f"{invoice_type}\t{data}\n"
        else:
            line = "\t".join(str(val) for val in row) + "\n"

        output_file.write(line)

    output_file.close()


# Подключаемся к SFTP-серверу
with pysftp.Connection(
    host=sftp_host,
    port=sftp_port,
    username=sftp_username,
    password=sftp_password
) as sftp:
    sftp.put('output.txt', remotepath=f'{sftp_directory}{excel_file.replace(".xlsx", ".txt")}')

print("Данные отправлены на сервер")
