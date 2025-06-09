
# from ROD_11_._9mm_label_ROD_11 import open_ptouch_insert_image9mm
# from ROD_11_._6mm_label_ROD_11 import funkcjonal6mmLong
# from ROD_11_._6mm_label_ROD_11_long import funkcjonal6mmLongMax
from ROD_11_._mm6_Logic import generator6mm_normal, napisy6mmZlaczkiOdp
from ROD_11_._mm6_Logic_long import generator6mm_long, napisy6mmZlaczkiOdp_long
from ROD_11_._mm9_Logic import generator9mm, napisy9mmZlaczkiOdp
from ROD_11_._24mm_Logic import odpalarka24mm
import sys, os

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
# sys.path.append(os.path.dirname(os.path.abspath(__file__)))
# sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ROD_11_'))


if __name__ == '__main__':
    import tkinter as tk
    import threading


    def klikaj_myszka():
        # Tutaj wybierz funkcję do uruchomienia:
        generator6mm_normal(napisy6mmZlaczkiOdp)
        generator6mm_long(napisy6mmZlaczkiOdp_long)
        generator9mm(napisy9mmZlaczkiOdp)
        odpalarka24mm()


    klikaj_myszka()


    #     # Po zakończeniu działania automatycznie zamknij GUI
    #     root.destroy()  # lub root.destroy()
    #
    # def start_klikania():
    #     threading.Thread(target=klikaj_myszka, daemon=True).start()
    #
    # def zatrzymaj_program(event=None):
    #     root.quit()  # lub root.destroy()

    # # GUI
    # root = tk.Tk()
    # root.title("Generator LBX")
    #
    # window_width = 250
    # window_height = 120
    # screen_width = root.winfo_screenwidth()
    # screen_height = root.winfo_screenheight()
    # x_position = screen_width - window_width - 10
    # y_position = 10
    # root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
    # root.attributes("-topmost", True)
    #
    # label = tk.Label(root, text="Praca w toku...", font=("Arial", 12))
    # label.pack(pady=10)
    #
    # stop_btn = tk.Button(root, text="STOP", command=zatrzymaj_program, bg="red", fg="white", font=("Arial", 14))
    # stop_btn.pack(pady=10)
    #
    # root.bind("<Escape>", zatrzymaj_program)
    #
    # start_klikania()
    # root.mainloop()