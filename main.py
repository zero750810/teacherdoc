try:
    import sys
    from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, 
                                QPushButton, QLabel, QVBoxLayout)
    import teacher_doc_generator
except ImportError as e:
    print(f"無法引入必要模組：{str(e)}")
    print("正在嘗試重新安裝 PyQt6...")
    import subprocess
    try:
        # 先移除現有的 PyQt6
        subprocess.run([sys.executable, "-m", "pip", "uninstall", "-y", "PyQt6", "PyQt6-Qt6", "PyQt6-sip"])
        # 重新安裝 PyQt6
        subprocess.run([sys.executable, "-m", "pip", "install", "--no-cache-dir", "PyQt6"])
        
        # 再次嘗試引入
        from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, 
                                   QPushButton, QLabel, QVBoxLayout)
    except Exception as install_error:
        print(f"安裝失敗：{str(install_error)}")
        sys.exit(1)

if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = teacher_doc_generator.TeacherDocApp()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"程式執行錯誤：{str(e)}")
        sys.exit(1) 