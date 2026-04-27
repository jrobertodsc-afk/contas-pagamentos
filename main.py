import customtkinter as ctk
from gui.app_main import ComSystemApp

if __name__ == '__main__':
    # Configuração global de aparência
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue") # Será sobrescrito pelo tema preto/lima
    
    app = ComSystemApp()
    app.mainloop()