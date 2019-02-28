import win32com.client
import os

from concurrent.futures import ThreadPoolExecutor

queue_info = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name = os.getenv('COMPUTERNAME')

queues = ['queue_test_1', 'queue_test_2']


def receive_messages(queue_name: str):

    queue_info.FormatName = f'direct=os:{computer_name}\\PRIVATE$\\{queue_name}'
    queue = None

    try:
        queue = queue_info.Open(1, 0)

        while True:
            msg = queue.Receive()
            print(f'Got Message from {queue_name}: {msg.Label} - {msg.Body}')

    except Exception as e:
        print(f'Error! {e}')

    finally:
        queue.Close()


def main():
    with ThreadPoolExecutor(max_workers=2) as executor:
        for queue in queues:
            executor.submit(receive_messages, queue)


if __name__ == '__main__':
    main()
