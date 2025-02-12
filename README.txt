*** File Readme này hướng dẫn cài đặt và chạy Chương trình test Jig.
*** Created on: 21-Jan-2025
*** Created by: Lê Hữu An (anlh55)

1. Cài đặt:

   - Tải và lưu Folder "App_Jig_v0" vào ổ đĩa D, trong Folder "App_Jig_v0" thực hiện cài đặt các phần mềm sau:

	a. Cài đặt python 3.10:

		B1: Mở folder Python_3_10
	
		B2: Click chạy python-3-10-1.exe

		B3: Thực hiện các thao tác install, lưu ý lưu đường dẫn vào file setup trong folder Python_3_10.

	b. Cài đặt STM32_CUBE_PROGRAMMER_CLI: dùng để nạp code cho mạch cần test

		B1: Mở folder STM32_CLI

		B2: Click chạy file exe để install phần mềm nạp code

		B3: Thực hiện các thao tác install, lưu ý lưu đường dẫn vào file setup trong folder STM32_CLI.

2. Tạo môi trường ảo và install thư viện:

	B1: Mở Terminal (nếu có) hoặc Powershell trên window.

	B2: Vào thư mục chứa folder App_Jig_v0 (thông thường nên lưu vào ổ đĩa D), bằng cách nhập vào:
	
		D:

	B3: Vào Folder App_Jig_v0 bằng câu lệnh:

		cd .\App_Jig_v0\
	
	B4: Chạy lệnh tạo môi trường ảo (myenv):
	
		D:\App_Jig_v0\Python_3_10\setup\bin -m venv myenv

	(** Nếu bị lỗi thì chạy lệnh để cấp quyền):

		Set-ExecutionPolicy RemoteSigned

	B5: Chờ quá trình tạo môi trường ảo hoàn tất, tạo Folder "myenv"


3. Chạy chương trình:

	+ B1: Mở Terminal (nếu có) hoặc Powershell trên window.

	+ B2: Vào thư mục chứa folder App_Jig_v0 (thông thường nên lưu vào ổ đĩa D), bằng cách nhập vào:
	
		D:

	+ B3: Vào Folder App_Jig_v0 bằng câu lệnh:

		cd .\App_Jig_v0\

	+ B4: Chạy môi trường ảo:
		
		.\myenv\Scripts\activate

	+ B5: Chạy chương trình test Jig bằng câu lệnh:

		python .\App_Jig.py

	+ B6: Bắt đầu chương trình test theo chỉ dẫn hiện hành trên cửa sổ terminal.

** Cách 2: sau khi chạy B3, chạy file "APP_JIG_Run.bat"

** Lưu ý: Trong quá trình chạy có lỗi chỉ cần nhấn Ctrl+C, sau đó thực hiện lại bước 5 là được.
		

