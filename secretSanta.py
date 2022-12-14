import csv
import random
import tkinter as tk
import sys
import win32com.client as win32
from tkinter import filedialog, messagebox


def read_csv():
    root = tk.Tk()
    root.withdraw()
    filename = filedialog.askopenfilename(
        title="Choose a csv file",
        filetypes=(
            ("CSV", "*.csv"),
            ("all files", "*.*s")
        )
    )
    root.destroy()
    with open(filename) as data:
        data = csv.reader(data)
        dict_from_csv = {rows[0]: rows[1] for rows in data}

    return dict_from_csv


def make_pairs(data):
    data = list(data.items())
    random.shuffle(data)
    return [
        (name, data[(i + 1) % len(data)])
        for i, name in enumerate(data)
    ]


def ask_send_confirmation():
    root = tk.Tk()
    root.withdraw()
    send_confirmation = messagebox.askyesno(
        title="Confirmation d'envoi",
        message="Voulez vous envoyer les résultats par mail ?"
    )
    root.destroy()
    if send_confirmation is False:
        print("Sending aborted !")
        sys.exit()
    else:
        return True


def send_results(pairs):
    ask_send_confirmation()
    outlook = win32.Dispatch('outlook.application')
    for pair in pairs:
        sender_name = pair[0][0].split()[0]
        sender_email = pair[0][1]
        receiver_name = pair[1][0].split()[0]
        mail = outlook.CreateItem(0)
        mail.Subject = "Secret Santa 2022"
        mail.To = sender_email
        mail.HTMLBody = r"""
        Beep...Beep...Boop...Bonjour {sender},<br><br>
        Je suis le robot lutin chargé de t'annoncer le résultat de ton tirage au sort.<br><br>
        Félicitations ! Tu es chargé(e) d'offrir un cadeau à {receiver}<br><br>
        Mais chut c'est un secret !<br><br>
        Sur ce, Joyeux Noël !!!
        <body>
            <pre>
           ____
          /    \
 ._.     /___/\ \
:(_):    |6.6| \|
  \\     '.-.'  O
   \\____.-"-.____
   '----|     |--.\
        |==[]=|  _\\_
         \___/    /|\
         // \\
        //   \\
        \\    \\ 
        _\\    \\__
       (___|    \__)
            </pre>
        </body>
        """.format(sender=sender_name, receiver=receiver_name)
        mail.Send()


if __name__ == "__main__":
    data = read_csv()
    pairs = make_pairs(data)
    send_results(pairs)
    print("mail all sent !")
