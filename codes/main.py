import os
from interface import Interface
from tools import Conciliation
import tkinter as tk


if __name__ == "__main__":
    root = tk.Tk()
    interface = Interface(
        root, path=os.path.join(os.path.dirname(__file__), "..", "result")
    )
    conciliation = Conciliation()
    conciliation.set_output(interface.get_output_folder())
    interface.set_action(lambda files: conciliation.new_conciliation(files))
    root.mainloop()
