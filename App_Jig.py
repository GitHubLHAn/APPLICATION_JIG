import serial
import time
import random

import serial.tools.list_ports

from uart import *
from colorama import Fore, Style
from openpyxl import Workbook

from loadcode import *
# from Scan_QR import scan_qr_from_camera

from led_func import Led_process
from sensor_func import Sensor_process
from center_func import Center_process
from boost_func import Boost_process
from power_func import Power_process

from log_data import log_test_result_func

import winsound

firm_path_goc = "D:\App_Jig_v0\Program_all_board"
log_path_goc = "D:\App_Jig_v0\Test_Result_Log"

print("======================================================================================")
print("                          BẮT ĐẦU CHƯƠNG TRÌNH TEST JIG")
print("======================================================================================\n")

# print("1. Mạch Boost")
# print("2. Mạch Center")
# print("3. Mạch Led")
# print("4. Mạch APR")
# print("5. Mạch Sensor")
# print("6. Mạch Power")
# print("7. Mạch Doc Charge")

Scan_Mach_Can_Test = ''
Header = ''
MachCanTest = ''
LoSanXuat = ''
Maso = ''

# Liệt kê các Cổng COM có sẵn
ports = serial.tools.list_ports.comports()
if not ports:
        print(Fore.RED+"-> Không tìm thấy cổng COM nào!\n"+ Style.RESET_ALL)
        result_COM = False
else:
    print(Fore.GREEN+"> Các cổng COM hiện có:"+ Style.RESET_ALL)
    for port in ports:
        print(f" - {port.device}: {port.description}")
    print("")
    port = "COM" + input(Fore.LIGHTWHITE_EX+"> Nhập cổng COM: "+ Style.RESET_ALL)
    # port = 'COM4'
    uart_var = uart()
    # Gọi hàm để bắt đầu nhận dữ liệu
    result_COM = uart_var.Init_COM(port)
    # print(result_COM)

while result_COM:
    #  Bắt đầu chương trình
    print(Fore.MAGENTA + "     **************************************************" + Style.RESET_ALL)
    print("     "+Fore.MAGENTA+"*"+Style.RESET_ALL+"   1. Quét mạch    "+Fore.MAGENTA+"*"+Style.RESET_ALL+"     2. Tắt chương trình    "+Fore.MAGENTA+"*"+Style.RESET_ALL)
    print(Fore.MAGENTA + "     **************************************************" + Style.RESET_ALL)
    
    lenhbatdau = input("> Nhập lệnh: ")
    
    while lenhbatdau != '1' and lenhbatdau != '2':
        print(Fore.RED + "-> Nhập sai lệnh, yêu cầu nhập lại!" + Style.RESET_ALL)
        lenhbatdau = input("> Nhập lệnh: ")
    
    if lenhbatdau == '1':
        print("\n-----------------------------------------------------------------------------------------------------------------------------------------------------------\n")
    elif lenhbatdau == '2':
        print("\n   <==> KẾT THÚC CHƯƠNG TRÌNH\n")
        break
    
    # Tạo một số ngẫu nhiên có 6 chữ số
    # random_serial = random.randint(100000, 999999)
    # Scan_Mach_Can_Test = "VTP_LED_111_" + str(random_serial)
    # Scan_Mach_Can_Test = "VTP_PWR_1234_457735"
    
    
    # Scan_Mach_Can_Test = scan_qr_from_camera()
    print(Fore.LIGHTWHITE_EX+"> Quét mã QR có trên mạch...\n"+ Style.RESET_ALL)

    check_infor = False
    while check_infor == False:
        #  Quét loại mạch cần test
        Scan_Mach_Can_Test = input(str("Scanning :  "))
        # Lấy thông tin
        # part_infor = Scan_Mach_Can_Test.split('_')
        if len(Scan_Mach_Can_Test.split('_')) == 4:
            Header, MachCanTest, LoSanXuat, Maso = Scan_Mach_Can_Test.split('_')
        
            if Header == "VTP" and MachCanTest in ['BS', 'CT', 'LED', 'AP', 'SS', 'PWR', 'DC']:
                check_infor = True
                print(Fore.YELLOW +"-> Header: " + Header + Style.RESET_ALL)
                print(Fore.YELLOW +"-> Loại mạch: " + MachCanTest + Style.RESET_ALL)
                print(Fore.YELLOW +"-> Lô Sản xuất: " + LoSanXuat + Style.RESET_ALL)
                print(Fore.YELLOW +"-> Mã số (Serial Number): " + Maso + Style.RESET_ALL)
                print("")
                print(Fore.YELLOW +"> Đã phát hiện mạch hợp lệ, yêu cầu đặt mạch vào Bộ JIG, sau đó nhập 1 để bắt đầu test!" + Style.RESET_ALL)
                while True:
                    lenhtest = input("> Nhập lệnh: ")
                    # lenhtest = input("> Nhập số Serial Number: \n")
                    if lenhtest == '1':
                    # if lenhtest == Maso:
                        break
                    else:
                        print(Fore.RED +"> Nhập sai! Yêu cầu nhập lại" + Style.RESET_ALL)
            else:
                print(Fore.RED +"> Mã QR không hợp lệ! Yêu cầu quét lại..." + Style.RESET_ALL)
        else:
                print(Fore.RED +"> Mã QR không hợp lệ! Yêu cầu quét lại..." + Style.RESET_ALL)
    
   
    
    # Test mạch BOOST:***************************************************************
    if MachCanTest == 'BS':
        print("                     BẮT ĐẦU TEST MẠCH BOOST                \n")
            
        ket_qua_test = Boost_process(uart_var, log_path_goc, Scan_Mach_Can_Test)
        in_ket_qua = f"===> KẾT QUẢ TEST MẠCH BOOST : {ket_qua_test}"
        print("---------------------------------------------------")
        print(Fore.YELLOW + in_ket_qua + Style.RESET_ALL)
        print("---------------------------------------------------\n")
        
    # Test mạch CENTER:**************************************************************
    elif MachCanTest == 'CT':   
        print("                     BẮT ĐẦU TEST MẠCH CENTER                \n")
        
        ket_qua_test = Center_process(uart_var, firm_path_goc, log_path_goc, Scan_Mach_Can_Test)
        in_ket_qua = f"===> KẾT QUẢ TEST MẠCH CENTER : {ket_qua_test}"
        print("---------------------------------------------------")
        print(Fore.YELLOW + in_ket_qua + Style.RESET_ALL)
        print("---------------------------------------------------\n")
      
    # Test mạch LED:*****************************************************************  
    elif MachCanTest == 'LED': 
        print("                    BẮT ĐẦU TEST MẠCH LED               \n")
        
        ket_qua_test =  Led_process(uart_var, firm_path_goc, log_path_goc, Scan_Mach_Can_Test)
        in_ket_qua = f"===> KẾT QUẢ TEST MẠCH LED : {ket_qua_test}"
        print("---------------------------------------------------")
        print(Fore.YELLOW + in_ket_qua + Style.RESET_ALL)
        print("---------------------------------------------------\n")
          
    # Test mạch APR:*****************************************************************  
    elif MachCanTest == 'APR': 
        print("-> BẮT ĐẦU TEST MẠCH APR\n")
        
    # Test mạch SENSOR:**************************************************************        
    elif MachCanTest == 'SS': 
        print("                     BẮT ĐẦU TEST MẠCH SENSOR                \n")
        
        ket_qua_test = Sensor_process(uart_var, firm_path_goc, log_path_goc, Scan_Mach_Can_Test)
        in_ket_qua = f"===> KẾT QUẢ TEST MẠCH SENSOR : {ket_qua_test}"
        print("---------------------------------------------------")
        print(Fore.YELLOW + in_ket_qua + Style.RESET_ALL)
        print("---------------------------------------------------\n")
        
    # Test mạch POWER:*****************************************************************    
    elif MachCanTest == 'PWR': 
        print("                     BẮT ĐẦU TEST MẠCH POWER                \n")
        
        ket_qua_test = Power_process(uart_var, firm_path_goc, log_path_goc, Scan_Mach_Can_Test)
        in_ket_qua = f"===> KẾT QUẢ TEST MẠCH POWER : {ket_qua_test}"
        print("---------------------------------------------------")
        print(Fore.YELLOW + in_ket_qua + Style.RESET_ALL)
        print("---------------------------------------------------\n")
            
    # Test mạch DOC CHARGE:************************************************************        
    elif MachCanTest == 'DC': 
                   
        print("-> BẮT ĐẦU TEST MẠCH DOC CHARGE\n")
        
        
    # *********************************************************************************        
    else:
        print(Fore.RED + "\n-> Mã QR không hợp lệ!\n" + Style.RESET_ALL)

    if MachCanTest != '':
        log_test_result_func(log_path_goc, Scan_Mach_Can_Test, ket_qua_test, MachCanTest)
     
    winsound.Beep(3000, 500)
    time.sleep(0.1)
    winsound.Beep(3000, 500)
    time.sleep(0.1)
    winsound.Beep(3000, 1000)    
    print(Fore.MAGENTA +"<-> Kết thúc test Jig, bạn có muốn test lại không?\n"+ Style.RESET_ALL)

if result_COM == True:
    # Đóng cổng COM
    uart_var.close_com()

