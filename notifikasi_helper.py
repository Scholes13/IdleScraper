import os
import sys

# Import untuk notifikasi desktop
try:
    from plyer import notification
    NOTIFICATIONS_AVAILABLE = True
    print("Plyer library imported successfully. Notifications are available.")
except ImportError:
    NOTIFICATIONS_AVAILABLE = False
    print("Notifikasi desktop tidak tersedia. Jalankan 'pip install plyer' untuk mengaktifkan fitur ini.")

def show_desktop_notification(title, message, app_name="Werkudara Maps Scraper", timeout=10, log_func=print):
    """
    Menampilkan notifikasi desktop jika library plyer tersedia
    
    Args:
        title (str): Judul notifikasi
        message (str): Isi pesan notifikasi
        app_name (str): Nama aplikasi yang ditampilkan
        timeout (int): Durasi notifikasi (dalam detik)
        log_func (function): Fungsi untuk mencatat log, default ke print
    
    Returns:
        bool: True jika notifikasi berhasil ditampilkan, False jika gagal
    """
    if not NOTIFICATIONS_AVAILABLE:
        log_func("Notifikasi desktop tidak tersedia. Pasang library 'plyer' untuk mengaktifkan fitur ini.")
        return False
    
    try:
        # Cari icon aplikasi
        icon_path = None
        possible_paths = [
            'app_icon.ico',
            'important_files/app_icon.ico',
            'assets/icons/app_icon.ico',
            'icon.ico'
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                icon_path = os.path.abspath(path)
                print(f"Found icon at: {icon_path}")
                break
        
        if icon_path is None:
            print("No icon found. Will use default system icon.")
        
        # Tampilkan notifikasi
        print(f"Attempting to show notification: {title} - {message}")
        notification.notify(
            title=title,
            message=message,
            app_name=app_name,
            timeout=timeout,
            app_icon=icon_path
        )
        
        log_func(f"Notifikasi desktop ditampilkan: {title}")
        return True
    except Exception as e:
        log_func(f"Gagal menampilkan notifikasi: {str(e)}")
        print(f"Exception details: {e.__class__.__name__}: {str(e)}")
        return False

# Test fungsi jika file dijalankan langsung
if __name__ == "__main__":
    print("Testing desktop notification...")
    result = show_desktop_notification(
        "Test Notifikasi", 
        "Ini adalah test notifikasi desktop dari Werkudara Maps Scraper."
    )
    print(f"Notification result: {result}")
    print("If result is True but you don't see a notification, check your system notification settings.") 