# MSMQ Producer and Consumer

Simple scripts to send and receive MSMQ messages

The script uses the `pywin32` library

To run:

1. Enable MSMQ on Windows Features
2. Create sample queues `compmgmt.msc` -> `Services and Applications` -> `Message Queueing` -> `Private Queues` -> `Right Click, New`
3. `pip install pywin32`
4. `python producer.py`
5. `python consumer.py` 