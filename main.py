# c:\Mac\Home\Documents\ACSOFT SENDER\main.py



def main():
    """
    Hàm chính để chạy ứng dụng. aa
    """
    print("Dự án Python mới đã được thiết lập trong VS Code!")
    try:
        response = requests.get("https://www.google.com", timeout=5)
        if response.status_code == 200:
            print("Đã kết nối Internet thành công.")
    except requests.ConnectionError:
        print("Không thể kết nối Internet.")

if __name__ == "__main__":
    main()
