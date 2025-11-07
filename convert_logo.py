import base64

# Thay 'logo.png' bằng tên file logo của bạn
file_name = 'logo.png'

try:
    with open(file_name, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read())

    # Tạo file output để dễ copy
    with open('logo_base64.txt', 'wb') as f_out:
        f_out.write(encoded_string)

    print(f"Thành công! Đã chuyển đổi '{file_name}' sang Base64.")
    print("Chuỗi Base64 đã được lưu vào file 'logo_base64.txt'.")
    print("Hãy mở file đó, sao chép toàn bộ nội dung và dán vào biến LOGO_PNG_BASE64 trong main.py.")

except FileNotFoundError:
    print(f"Lỗi: Không tìm thấy file '{file_name}'. Vui lòng đặt file này cùng thư mục.")
except Exception as e:
    print(f"Đã xảy ra lỗi: {e}")

