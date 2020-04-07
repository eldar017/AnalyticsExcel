from openpyxl import Workbook
import json
import pika
from openpyxl import load_workbook

credentials = pika.PlainCredentials(username='rabbit', password='rabbit')
connection = pika.BlockingConnection(
    pika.ConnectionParameters(host='192.168.146.134', port='5672', credentials=credentials))
channel = connection.channel()

channel.queue_declare(queue='hello')
counter = 1

def callback(ch, method, properties, body):
    print(" [x] Received %r" % body)
    data = str(body).split(",")
    userID= data[2]
    #print(userID)
    excel = body
    excel_dict = json.loads(excel)
    #ecoded_response = response.read().decode("UTF-8")
    print(json.dumps(excel_dict,indent=4,sort_keys=True))
    excel_raw = excel_dict["xlData"]
    #print (excel_raw)
    _wb = load_workbook(excel_raw)
    _wb1 = _wb.active
    print(_wb1)



channel.basic_consume(
    queue='hello', on_message_callback=callback, auto_ack=True)

print(' [*] Waiting for messages. To exit press CTRL+C')
channel.start_consuming()