

from ROD_11_._mm6_Logic import generator6mm_normal, napisy6mmZlaczkiOdp
from ROD_11_._mm6_Logic_long import generator6mm_long, napisy6mmZlaczkiOdp_long
from ROD_11_._mm9_Logic import generator9mm, napisy9mmZlaczkiOdp
from tamuryn.klas16 import odpalarka24bis
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
        odpalarka24bis()


    klikaj_myszka()
