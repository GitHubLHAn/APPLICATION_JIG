import subprocess
import sys



def flash_firmware(firmware_path, port="SWD"):
    try:
        # command = [r"D:\Program Storage\CUBE_PROG__\setup\bin\STM32_Programmer_CLI.exe", "-c", "port=SWD", "-disconnect"]
        # result = subprocess.run(command, capture_output=True, text=True)
        
        # Cấu hình lệnh CLI
        command = [
            r"D:\Program Storage\CUBE_PROG__\setup\bin\STM32_Programmer_CLI.exe",
            "-c", f"port={port}",
            "-w", firmware_path,
            "-v",
            "-rst"
        ]

        # Gọi lệnh
        result = subprocess.run(command, capture_output=True, text=True)
        
        # Hiển thị kết quả
        # print("Kết quả:", result.stdout)
        # print("Lỗi:", result.stderr)
        
        # Kiểm tra trạng thái
        if result.returncode == 0:
            # print("Nạp firmware thành công!")
            return 1
        else:
            # print("Lỗi khi nạp firmware.")
            return -1

    except FileNotFoundError:
        print("Không tìm thấy STM32_Programmer_CLI. Hãy kiểm tra lại đường dẫn.")
    except Exception as e:
        print(f"Lỗi xảy ra: {e}")

# ---------------------------------------------------------------------------------------

firm_path_goc = "D:\App_Jig_v0\Program_all_board"

print("Bat dau nap!")
# Đường dẫn tới file firmware
firmware_path = firm_path_goc + "\power_board\power_main.hex"

# firmware_path = "D:\App_Jig_v0\Program_all_board\power_board\power_test.hex"

while True:
    # firmware_path = input("Nhap duong dan file .hex: ")

    # Nạp firmware
    kq = flash_firmware(firmware_path)
    
    if kq == 1:
        print("Nạp firmware thành công!")
    else:
        print("Lỗi khi nạp firmware.")

    print("Ban co muon chay lai chuong trinh khong:")
    print("1. Co")
    print("2. Khong")
    i = input()
    
    if i == '2':
        print("Ket thuc chuong trinh")
        break

# D:\Test_nap_code_CLI_cube\Program_all_board\center_board\center_test.hex