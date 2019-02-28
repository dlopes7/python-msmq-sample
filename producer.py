import win32com.client
import os
import time
import random

queue_info = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name = os.getenv('COMPUTERNAME')

queues = ['queue_test_1', 'queue_test_2']


def send_message(queue_name: str, label: str, message: str):

    queue_info.FormatName = f'direct=os:{computer_name}\\PRIVATE$\\{queue_name}'
    queue = None

    try:
        queue = queue_info.Open(2, 0)

        msg = win32com.client.Dispatch("MSMQ.MSMQMessage")
        msg.Label = label
        msg.Body = message

        msg.Send(queue)

    except Exception as e:
        print(f'Error! {e}')

    finally:
        queue.Close()


def main():
    i = 0
    while True:
        i += 1
        send_message(random.choice(queues), 'test label', f'{i}: this is a test message')
        print(f'{i}: Message sent!')
        time.sleep(0.5)


if __name__ == '__main__':
    main()
