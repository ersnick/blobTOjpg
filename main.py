import mysql.connector
from PIL import Image
import io
from openpyxl import load_workbook


# конфиг бд
mydb = mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="perenos"
)

# Создаем курсор
cursor = mydb.cursor()

# пишем свой запрос
query = 'SELECT CODE_PPL, PHOTO_PPL, CONCAT(LNAME_PPL, " ", FNAME_PPL, " ", PNAME_PPL), MPHONE_PPL FROM PEOPLE'

# результат запроса
cursor.execute(query)
data = cursor.fetchall()

for record in data:

    try:

        # тут очередность полей как в запросе
        id_people = record[0]
        image_data = record[1]
        fio = record[2]
        phone = record[3]

        # конвертируем изображение
        image = Image.open(io.BytesIO(image_data))

        # Convert RGBA to RGB if the image has an alpha channel
        if image.mode == 'RGBA':
            image = image.convert('RGB')

        # сохраняем как jpg
        image.save(f'photos/{id_people}.jpg', "JPEG")

        # записываем в эксель
        wb = load_workbook('import_photo.xlsx')
        ws = wb.active
        dataForXlsx = [id_people, f'{id_people}.jpg', fio, phone]
        ws.append(dataForXlsx)

        # сохраняем эксельку
        wb.save('import_photo.xlsx')

    # на случай если картинки нет
    except Exception as e:
        print(f"{e}\nСкорее всего у это пользователя просто фотки нет, его id: {id_people}")

