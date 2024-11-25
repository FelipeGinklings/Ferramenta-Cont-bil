import os
import tkinter as tk
from tkinter import filedialog


class Interface:
    def __init__(self, root, path):
        self._root = root
        self._action = lambda: None
        self._output_folder = path
        self._bg_color = "#333333"
        self._top_frame_color = "#123524"
        self._button_bg_color = "#123524"
        self._button_fg_color = "#FFFFFF"
        self._button_active_bg_color = "#184c33"
        self._label_bg_color = "#222222"
        self._label_fg_color = "#FFFFFF"
        self._bar_color = "#000000"

        self._ensure_result_folder()
        self._root.title("Retro Reconciliation Interface")
        self._root.geometry("480x400")  # Ajuste a largura da tela
        self._root.configure(bg=self._bg_color)  # Dark background for retro style
        self._root.resizable(False, True)  # Trava a largura da janela
        self._create_widgets()

    def _ensure_result_folder(self):
        result_folder = os.path.join(os.path.dirname(__file__), "..", "result")
        if not os.path.exists(result_folder):
            os.makedirs(result_folder)
        self._output_folder = result_folder

    def _create_widgets(self):
        self._create_top_frame()
        self._create_separator()  # Adiciona a linha preta entre os frames
        self._create_files_frame()

    def _create_separator(self):
        separator = tk.Frame(self._root, bg=self._bar_color, height=2)
        separator.pack(fill="x")

    def _action_new(self):
        # Filter to show only .csv
        files = filedialog.askopenfilenames(
            title="Selecione um arquivo CSV",
            filetypes=(("Arquivos CSV", "*.csv"), ("Todos os Arquivos", "*.*")),
        )
        self._action(files)
        self._update_folder_view()

    def _open_about(self):
        about_window = tk.Toplevel(self._root)
        about_window.title("About")
        about_window.geometry("400x250")
        about_window.configure(bg=self._bg_color)  # Dark background for retro style

        text = (
            "This program performs financial data reconciliation, "
            "allowing you to check for discrepancies between records. "
            "\nFolder: Folder where to look for files "
            "\nNew: to start a new reconciliation. "
        )

        # Explanatory text with retro style
        text_widget = tk.Text(
            about_window,
            bg=self._label_bg_color,
            fg=self._label_fg_color,
            wrap="word",
            font=("Courier New", 12),
            relief="flat",
            height=8,
            width=50,
        )
        text_widget.pack(padx=10, pady=10, fill="both", expand=True)

        text_widget.insert("1.0", text)
        text_widget.tag_add("bold", "2.0", "2.7")  # 'Folder:'
        text_widget.tag_add("bold", "3.0", "3.4")  # 'New:'
        text_widget.tag_configure("bold", font=("Courier New", 12, "bold"))
        text_widget.configure(state="disabled")

        # Close button
        close_button = tk.Button(
            about_window,
            text="Close",
            command=about_window.destroy,
            bg=self._button_bg_color,
            fg=self._button_fg_color,
            font=("Courier New", 12),
            relief="flat",
            activebackground=self._button_active_bg_color,
        )
        close_button.pack(pady=12)

    def _create_top_frame(self):
        top_frame = tk.Frame(self._root, bg=self._top_frame_color)  # Uniform color line
        top_frame.pack(fill="x")

        # "New" button
        self.new_button = tk.Button(
            top_frame,
            text="New",
            command=self._action_new,
            bg=self._button_bg_color,
            fg=self._button_fg_color,
            font=("Courier New", 12),
            relief="flat",
            activebackground=self._button_active_bg_color,
            width=15,
        )
        self.new_button.grid(row=0, column=2, pady=0)

        # Black bar between "New" and "Folder"
        black_bar2 = tk.Frame(top_frame, bg=self._bar_color, width=2, height=30)
        black_bar2.grid(row=0, column=3, pady=0)

        # "Folder" button
        self.folder_button = tk.Button(
            top_frame,
            text="Folder",
            command=self._select_folder,
            bg=self._button_bg_color,
            fg=self._button_fg_color,
            font=("Courier New", 12),
            relief="flat",
            activebackground=self._button_active_bg_color,
            width=15,
        )
        self.folder_button.grid(row=0, column=4, pady=0)

        # Black bar between "Folder" and "About"
        black_bar3 = tk.Frame(top_frame, bg=self._bar_color, width=2, height=30)
        black_bar3.grid(row=0, column=5, pady=0)

        # "About" button
        self.about_button = tk.Button(
            top_frame,
            text="About",
            command=self._open_about,
            bg=self._button_bg_color,
            fg=self._button_fg_color,
            font=("Courier New", 12),
            relief="flat",
            activebackground=self._button_active_bg_color,
            width=15,
        )
        self.about_button.grid(row=0, column=6, pady=0)

        # Black bar at the end of "About"
        black_bar4 = tk.Frame(top_frame, bg=self._bar_color, width=2, height=30)
        black_bar4.grid(row=0, column=7, pady=0)

    def _create_files_frame(self):
        form_frame = tk.Frame(self._root, bg=self._bg_color)
        form_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self._folder_view = tk.Frame(form_frame, bg=self._label_bg_color)
        self._folder_view.pack(fill="both", expand=True)
        self._update_folder_view()

    def _update_folder_view(self):
        for widget in self._folder_view.winfo_children():
            widget.destroy()
        for idx, item in enumerate(os.listdir(self._output_folder)):
            if item.endswith(".xlsx"):  # Filter to show only .xlsx files
                item_name = os.path.splitext(item)[0]
                row = idx

                file_label = tk.Label(
                    self._folder_view,
                    text=item_name,
                    bg=self._label_bg_color,
                    fg=self._label_fg_color,
                    font=("Courier New", 12),
                    width=25,
                    anchor="w",
                )
                file_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")

                open_button = tk.Button(
                    self._folder_view,
                    text="Open",
                    command=lambda item=item: self._open_file(item),
                    bg=self._button_bg_color,
                    fg=self._button_fg_color,
                    font=("Courier New", 10),
                    relief="flat",
                    activebackground=self._button_active_bg_color,
                )
                open_button.grid(row=row, column=1, padx=5, pady=5)

                rename_button = tk.Button(
                    self._folder_view,
                    text="Rename",
                    command=lambda item=item: self._rename_file(item),
                    bg=self._button_bg_color,
                    fg=self._button_fg_color,
                    font=("Courier New", 10),
                    relief="flat",
                    activebackground=self._button_active_bg_color,
                )
                rename_button.grid(row=row, column=2, padx=5, pady=5)

                delete_button = tk.Button(
                    self._folder_view,
                    text="Delete",
                    command=lambda item=item: self._delete_file(item),
                    bg=self._button_bg_color,
                    fg=self._button_fg_color,
                    font=("Courier New", 10),
                    relief="flat",
                    activebackground=self._button_active_bg_color,
                )
                delete_button.grid(row=row, column=3, padx=5, pady=5)

    def _open_file(self, item):
        file_path = os.path.join(self._output_folder, item)
        os.startfile(file_path)

    def _rename_file(self, item):
        new_name = filedialog.asksaveasfilename(
            initialdir=self._output_folder,
            initialfile=item,
            title="Rename file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if new_name:
            new_name = os.path.splitext(new_name)[0] + ".xlsx"
            os.rename(os.path.join(self._output_folder, item), new_name)
            self._update_folder_view()

    def _delete_file(self, item):
        os.remove(os.path.join(self._output_folder, item))
        self._update_folder_view()

    def _select_folder(self):
        self._output_folder = filedialog.askdirectory(title="Select a folder")
        if self._output_folder:
            self._update_folder_view()

    def _update_widgets_colors(self):
        # Atualiza as cores dos widgets existentes
        for widget in self._root.winfo_children():
            if isinstance(widget, tk.Frame):
                widget.configure(bg=self._bg_color)
            elif isinstance(widget, tk.Button):
                widget.configure(
                    bg=self._button_bg_color,
                    fg=self._button_fg_color,
                    activebackground=self._button_active_bg_color,
                )
            elif isinstance(widget, tk.Label):
                widget.configure(bg=self._label_bg_color, fg=self._label_fg_color)
        self._update_folder_view()

    def set_action(self, func):
        """
        Define a função de ação a ser executada quando um novo arquivo é selecionado.

        Parâmetros:
        func (callable): Função a ser chamada com a lista de arquivos selecionados.
        """
        self._action = func

    def get_output_folder(self):
        """
        Retorna o caminho da pasta de saída onde os arquivos processados são armazenados.

        Retorno:
        str: Caminho da pasta de saída.
        """
        return self._output_folder

    def set_colors(
        self,
        bg_color=None,
        top_frame_color=None,
        button_bg_color=None,
        button_fg_color=None,
        button_active_bg_color=None,
        label_bg_color=None,
        label_fg_color=None,
        bar_color=None,
    ):
        """
        Define as cores da interface.

        Parâmetros:
        bg_color (str): Cor de fundo da interface.
        top_frame_color (str): Cor do frame superior.
        button_bg_color (str): Cor de fundo dos botões.
        button_fg_color (str): Cor do texto dos botões.
        button_active_bg_color (str): Cor de fundo dos botões quando ativos.
        label_bg_color (str): Cor de fundo dos rótulos.
        label_fg_color (str): Cor do texto dos rótulos.
        bar_color (str): Cor das barras de separação.
        """
        if bg_color:
            self._bg_color = bg_color
        if top_frame_color:
            self._top_frame_color = top_frame_color
        if button_bg_color:
            self._button_bg_color = button_bg_color
        if button_fg_color:
            self._button_fg_color = button_fg_color
        if button_active_bg_color:
            self._button_active_bg_color = button_active_bg_color
        if label_bg_color:
            self._label_bg_color = label_bg_color
        if label_fg_color:
            self._label_fg_color = label_fg_color
        if bar_color:
            self._bar_color = bar_color
        self._root.configure(bg=self._bg_color)
        self._update_widgets_colors()


if __name__ == "__main__":
    root = tk.Tk()
    interface = Interface(
        root, path=os.path.join(os.path.dirname(__file__), "..", "result")
    )

    def example_action(files):
        print("Arquivos selecionados:", files)

    interface.set_action(example_action)
    root.mainloop()
