from serial import *
import serial
from time import sleep
from datetime import datetime

class SyringePump():
    def __init__(self, port, name = 'UNKNOWN (Syringe Pump)'):
        '''
        port: str
            port number, e.g. COM3
        '''
        
        self.con = serial.Serial(port)
        
        self.name = name 

        print('{} at {}'.format(self.name, self.con.port))

    def __repr__(self) -> str:
        return 'Syringe Pump'

    def start(self):
        command = 'start'
        arg = bytes(str(command), 'utf8') + b'\r'

        self.con.write(arg)

        action = '{}:\t\tStarted.'.format(self.name)
        print(action)
    
    def stop(self):
        command = 'stop'
        arg = bytes(str(command), 'utf8') + b'\r'

        self.con.write(arg)

        action = '{}:\t\tStopped.'.format(self.name)
        print(action)

    def pause(self):
        command = 'pause'
        arg = bytes(str(command), 'utf8') + b'\r'

        self.con.write(arg)

        action = '{}:\t\tPaused.'.format(self.name)
        print(action)

    def changeFlowrate(self, flowrate, start = True):
        command = 'set rate ' + str(flowrate) 

        arg = bytes(str(command), 'utf8') + b'\r'
        self.con.write(arg)
        
        today = datetime.now()
        CurrentTime = today.strftime("%H:%M:%S")

        action = '{}:\t\tFlowrate changed to {} ml/min at {}.'.format(self.name,flowrate,CurrentTime)
        print(action)
        if not start:
            sleep(0.2)
            self.pause()


     