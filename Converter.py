import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from tkinter import Tk, filedialog, simpledialog, messagebox
#from pandasgui import show

def generate_pdf(data, output_path):
    try:
        pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
    except IOError:
        messagebox.showerror("Błąd", "Nie znaleziono pliku czcionki Arial.ttf. Spróbuj użyć innej czcionki.")
        return

    pdf = canvas.Canvas(output_path)
    pdf.setFont('Arial', 12)

    labels = [
        "1. Nazwa jednostki:",
        "2. Numer zamówienia/umowy:",
        "3. Konto alokacji (sfinansowane ze środków):",
        "4. Miejsce Powstania Kosztu (Jednostka organizacyjna): ",
        "5. Źródło finansowania (Płacone z budżetu jedn.organizacyjnej):",
        "6. Projekt (zakup dla proj./grant/zadanie/zlecenie):",
        "7. Potwierdzenie kompletności dostawy/usług i zapisów na fakturze:",
        "8. Zatwierdzenie przez osobę upoważnioną przez Rektora PP:",
        "9. Podstawa prawna (Art. 4 pkt 8 PZP) udzielenia zamówienia:",
        "10. Przedmiot zamówienia:"
    ]

    x, y = 100, 700
    max_line_length = 50  # Maksymalna długość tekstu w jednej linii

    for label, value in zip(labels, data):
        full_text = f"{label} {value}"
        words = full_text.split()
        lines = [' '.join(words[i:i+2]) for i in range(0, len(words), 2)]

        current_line = lines[0]
        for line in lines[1:]:
            if len(current_line) + len(line) <= max_line_length:
                current_line += " " + line
            else:
                pdf.drawString(x, y, current_line)
                y -= 20
                current_line = line

        pdf.drawString(x, y, current_line)
        y -= 20

    pdf.save()

def read_csv():
    Tk().withdraw()  
    file_path = filedialog.askopenfilename(filetypes=[("Pliki CSV", "*.csv")])
    if not file_path:
        messagebox.showerror("Błąd", "Nie wybrano pliku CSV.")
        return None
    df = pd.read_csv(file_path, sep=';')

    # Wyświetlanie informacji o tabeli - pandasgui (OPTIONAL)
    #show(df, "Wczytany plik CSV")

    return df

def choose_output_path():
    Tk().withdraw()  
    file_path = filedialog.asksaveasfilename(filetypes=[("Pliki PDF", "*.pdf")])
    if not file_path:
        messagebox.showerror("Błąd", "Nie wybrano miejsca zapisu pliku PDF.")
        return None
    return file_path

def get_row_number():
    root = Tk()
    root.withdraw()  
    row_number = simpledialog.askinteger("Numer wiersza", "Podaj numer wiersza:")
    root.destroy()  
    return row_number

def main():
    df = read_csv()
    if df is None:
        return

    row_number = get_row_number()
    if row_number is None:
        return

    try:
        selected_row = df.iloc[row_number].tolist()
    except IndexError:
        messagebox.showerror("Błąd", "Wybrano nieprawidłowy numer wiersza.")
        return

    output_path = choose_output_path()
    if output_path is None:
        return

    generate_pdf(selected_row, output_path)
    messagebox.showinfo("Sukces", f"Dokument PDF został pomyślnie utworzony i zapisany w: {output_path}")

if __name__ == "__main__":
    main()
