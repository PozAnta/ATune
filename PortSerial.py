import serial
import time

serial.Serial()

def port(comnd, nameCom, Tx):

    ser = serial.Serial(nameCom, 115200, timeout=2, parity=serial.PARITY_NONE)  # open serial port

    for attempt in range(0, 5, 1):

        # print(ser.name)   # check which port was really used
        st1 = comnd
        st2 = "\r\n"
        st = st1 + st2

        ser.write(bytes(st, "ascii"))
        ch = ''
        start = time.time()
        timeout = 10
        last2ch = ''
        cmd = ''

        while not ((last2ch) == '->'):
            if (time.time() - start) > timeout: return "timeout"
            try:
                x = ser.read().decode("ascii")
            except UnicodeDecodeError:
                if attempt == 5:
                    print("R/W wrong data")
                    return "R/W wrong data"
                pass

            chp = ch
            ch = x
            last2ch = chp+ch
            cmd = cmd+x

        cmd1 = cmd[:len(cmd)-3] # delete "-->"
        if Tx == True:
            return cmd1
        else:
            cmd1 = cmd1.replace(st1, '')
            return cmd1
