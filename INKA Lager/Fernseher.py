try:
    from Tkinter import *
except:
    from tkinter import *

import openpyxl

"""Argen IS.exe - die Suchsoftware, um schnell die gewünschte Zelle zu finden. 
Die Software basiert auf der Programmiersprache Python. Wenn ein Artikel gescannt wird, 
wird dem Benutzer die Artikelnummer und der Platz auf dem Bildschirm angezeigt. Dadurch 
lassen sich die Waren (Abatement, Schrauben, usw.) schneller platzieren. Die Informationen 
stammen aus einer Excel-Tabelle auf dem Server."""


"""die Datei öffnen und lesen, um die Daten als "Liste von Listen" in den Arbeitsspeicher zu schreiben"""

book = openpyxl.open(r'\\srv-file\Auftrag\IT\ArgenIS\Artikel.xlsx', read_only=True)

sheet = book.active

l_gen, l_sub = [], []

for rows in sheet.iter_rows(min_row=2, min_col=0):
    l_sub = []
    for column in rows:
        l_sub.append(column.value)
    l_gen.append(l_sub)
print(l_gen)
book.close()

"""Erstellen UI(User-Interface)"""

win = Tk()
win.geometry(f'1920x1080')
win.title('Argen Output')
win.config(bg='black')
win.attributes('-fullscreen', True)


"""Die Funktion, um Software zu schließen"""

def quit(event):
    win.destroy()

"""Hauptfunktion"""

def get(event):

    """Definieren Globale-varible"""

    global label_oben_links, label_oben_rechts, label_mitte_gross
    label_oben_links.grid_forget()
    label_oben_rechts.grid_forget()
    label_mitte_gross.grid_forget()

    """Scannen den Artikel"""

    id_artikel = enter_artikel.get()[2:16]

    count = 0

    """Artikel und position suchen"""

    for i in l_gen:
        if id_artikel in i and i[3] == 'Vorne':
            label_oben_links.config(text=i[1])
            label_oben_rechts.config(text=i[3], fg='green')
            label_mitte_gross.config(text=i[2], fg='green', font=('Calibri', 500, 'bold'))
            count += 1
            break

        elif id_artikel in i and i[3] == 'Hinten':
            label_oben_links.config(text=i[1])
            label_oben_rechts.config(text=i[3], fg='orange')
            label_mitte_gross.config(text=i[2], fg='orange', font=('Calibri', 500, 'bold'))
            count += 1
            break

    if count == 0:
        for i in label_oben_links, label_oben_rechts:
            i.config(text='')
        label_mitte_gross.config(text='Artikel nicht gefunden', fg='red', font=('Calibri', 135, 'bold'))
        label_mitte_gross.grid_configure(pady=90)

    enter_artikel.delete(0, 'end')

    label_oben_links.grid(row=0, column=0)
    label_oben_rechts.grid(row=0, column=1)
    label_mitte_gross.grid(row=1, column=0, columnspan=2)

"""Erstellen und platzieren die Elemente auf dem Bildschirm"""

label_oben_links = Label(win, text='', fg="white", bg="black", font=('Calibri', 150, 'bold'))
label_oben_links.grid(row=0, column=0)

label_oben_rechts = Label(win, text='', bg="black", font=('Calibri', 150, 'bold'))
label_oben_rechts.grid(row=0, column=1)

label_mitte_gross = Label(win, text='Scannen Sie bitte den Artikel', anchor='e', fg="silver", bg="black",
                          font=('Calibri', 90, 'bold'))
label_mitte_gross.grid(row=1, column=0, columnspan=2, pady=150)

enter_artikel = Entry(win, bg='black', fg='green', font=('Calibri', 11, 'bold'))
enter_artikel.grid(row=3, column=0, columnspan=2, pady=800)
enter_artikel.bind('<Return>', get)
enter_artikel.focus_set()

"""Taste Escape = Exit"""

win.bind('<Escape>', quit)

win.grid_columnconfigure(0, minsize=960)
win.grid_columnconfigure(1, minsize=960)

win.mainloop()
