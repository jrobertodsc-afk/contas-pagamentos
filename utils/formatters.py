import tkinter as tk

__all__ = ['_aplicar_mascara_data', '_aplicar_mascara_valor']

def _aplicar_mascara_data(widget):
    """Aplica máscara de data (DD/MM/AAAA) a um widget Entry do Tkinter."""
    def format_date(event):
        if event.keysym in ("BackSpace", "Delete", "Left", "Right", "Tab"):
            return
        
        text = widget.get().replace('/', '')
        text = ''.join(c for c in text if c.isdigit())[:8]
        
        new_text = ""
        if len(text) > 0:
            new_text += text[:2]
        if len(text) > 2:
            new_text += '/' + text[2:4]
        if len(text) > 4:
            new_text += '/' + text[4:8]
            
        widget.delete(0, tk.END)
        widget.insert(0, new_text)

    widget.bind("<KeyRelease>", format_date)

def _aplicar_mascara_valor(entry):
    """Aplica máscara de valor monetário (R$ 1.234,56) em tempo real."""
    def format_money(event):
        if event.keysym in ("BackSpace", "Delete", "Left", "Right", "Tab"):
            return
        
        # Pega apenas dígitos
        text = entry.get()
        digits = "".join(re.findall(r"\d", text))
        if not digits:
            return
            
        try:
            # Converte centavos para decimal
            val = float(digits) / 100
            formatted = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            entry.delete(0, tk.END)
            entry.insert(0, formatted)
        except:
            pass

    import re
    entry.bind("<KeyRelease>", format_money)
