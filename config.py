import os.path
import datetime


if datetime.datetime.now().hour < 12:
    BODY = 'Доброе утро!\n\nВыгрузки с конкурсами во вложении\n\n\nТендерный отдел'
else:
    BODY = 'Добрый день!\n\nВыгрузки с конкурсами во вложении.\n\n\nТендерный отдел'    
COPY = "tenders@a1tis.ru"
