import json
import pika
from openpyxl import load_workbook
import pandas as pd
# credentials = pika.PlainCredentials(username='user', password='password')
# connection = pika.BlockingConnection(
#     pika.ConnectionParameters(host='54.144.236.23', port='5672', credentials=credentials))
# channel = connection.channel()
#
# channel.queue_declare(queue='hello',durable=False)
# fileName = "C:\OM\exampleMasterCard2.xlsx"
# wb = load_workbook(fileName)
# wb1 = wb.active
# data = pd.read_excel(fileName)
# d = str(data)
# dd = json.dumps(d,indent=4,sort_keys=True)
# #print(data)
# channel.basic_publish(exchange='Eldar',
#                           routing_key='eldar',
#                           body=dd)


link = 'https://s5.aconvert.com/convert/p3r68-cdx67/o82qi-2jxzt.xlsx'
d1 = pd.read_excel(link,'o82qi-2jxzt')
print(d1)