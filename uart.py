import serial

from colorama import Fore, Style

class uart:
    def Init_COM(self, _port):
        try:

            self.ser = serial.Serial(
                        port = _port,
                        baudrate=9600,
                        bytesize=serial.EIGHTBITS,
                        parity=serial.PARITY_NONE,
                        stopbits=serial.STOPBITS_ONE,
                        timeout= 1
                        )
            kq = f"-> Mở cổng {_port} thành công (9600)\n"
            print(Fore.GREEN + kq + Style.RESET_ALL)
            result = True
            
        except serial.SerialException as e:
            kq = "Mở cổng COM thất bại"
            print(Fore.RED + kq + Style.RESET_ALL)
            result = False
        return result

    def send_data(self,hex_array):
        self.frame_0 = hex_array[0].to_bytes(1, 'big')
        self.frame_1 = hex_array[1].to_bytes(1, 'big')
        self.frame_2 = hex_array[2].to_bytes(1, 'big')
        self.frame_3 = hex_array[3].to_bytes(1, 'big')
        self.frame_4 = hex_array[4].to_bytes(1, 'big')
        self.frame_5 = hex_array[5].to_bytes(1, 'big')
        self.frame_6 = hex_array[6].to_bytes(1, 'big')
        self.frame_7 = hex_array[7].to_bytes(1, 'big')
        self.frame_8 = hex_array[8].to_bytes(1, 'big')
        self.frame_9 = hex_array[9].to_bytes(1, 'big')
        self.frame_10 = hex_array[10].to_bytes(1, 'big')
        self.frame_11 = hex_array[11].to_bytes(1, 'big')
        
   
        self.ser.write(self.frame_0)
        self.ser.write(self.frame_1)
        self.ser.write(self.frame_2)
        self.ser.write(self.frame_3)
        self.ser.write(self.frame_4)
        self.ser.write(self.frame_5)
        self.ser.write(self.frame_6)
        self.ser.write(self.frame_7)
        self.ser.write(self.frame_8)
        self.ser.write(self.frame_9)
        self.ser.write(self.frame_10)
        self.ser.write(self.frame_11)
    def read_data(self):
        data_rec = self.ser.read(12)
        return data_rec

    def close_com(self):
        self.ser.close()






# def Init_COM(port, baudrate=9600, timeout=1):
#     try:
#         # Mở cổng UART
#         serial.Serial(port, baudrate, timeout=timeout)
#         print(f"-> Mở cổng {port} thành công ({baudrate})\n")

#     except serial.SerialException as e:
#         print("Mở cổng COM thất bại")



# def Send_CMD_to_JIG(port, hex_array):
#     try:
#         ser = serial.Serial(port, 9600, timeout=1)
#         data_to_send = bytearray(hex_array)
#         ser.write(data_to_send)  
#         print("sent") 
#     except serial.SerialException as e:
#         print(f"Lỗi mở cổng COM !!!")
    
    
# def Receive_RES_from_Jig(data_array, num_bytes, timeout):
#     while True:
#         start_time = time.time()
#         data = ser.read(num_bytes)
#         if data:
#             data_array.extend(data)
#             break
#         if(time.time() - start_time > timeout):
#             return -1
            
    
    