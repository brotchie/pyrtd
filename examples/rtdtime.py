"""
Example usage of pyrtd.RTDClient connecting to the RTDTime.RTD component. The
RTDTime.dll implementing this component is part of the "Building a Real-Time
Data Server in Excel 2002" MSDN article and is available at:

http://download.microsoft.com/download/4/9/c/49cb54f8-63b6-4024-845b-fd2c8b0d8917/odc_xlrtdbuild.exe

You'll have to register the RTD component by executing "regsvr32 RTDTime.dll".
This component sends an UpdateNotify every second, thus this example prints out
the current time every second.

"""
import pythoncom
from rtd import RTDClient

if __name__ == '__main__':
    time = RTDClient('RTDTime.RTD')
    time.connect()
    time.register_topic('Now')

    while 1:
        # This line is critical, it tells the pythoncom subsystem to
        # handle any pending windows messages. We're waiting for an
        # UpdateNotify callback from the RTDServer; if we don't
        # check for messages we'll never be notified of pending
        # RTD updates!
        pythoncom.PumpWaitingMessages()

        if time.update():
            print time.get('Now')
