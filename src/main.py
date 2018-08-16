from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from report_generator import *
import os


def main():
    c = Controller()
    c.run()


class Controller:
    def __init__(self):
        self._root = Tk()
        self._root.resizable(False, False)
        self._gui = ReportGUI(self._root)
        self._gui.generate_button.configure(command=self._generate_button_click)
        self._gui.open_button.configure(command=self._open_button_click)

        self._model = None

    def run(self):
        self._root.mainloop()

    def _generate_button_click(self):
        if self._validate_input_dir():
            output_directory = self._gui.out_directory_entry.get().rstrip('\\')
            self._model = ReportGenerator(output_directory)
            try:
                self._gui.destroy_generate_button()
                self._gui.show_generating_label()
                self._gui.root.update()
                self._model.generate()
            except ReportGeneratorException as err:
                self._gui.draw_error_message_box(1, 'Error', 'Error in generating report. Error: ' + str(err))
                self._rewind_gui()
            except Exception as err:
                self._rewind_gui()
                self._gui.draw_error_message_box(1, 'Error', 'Unexpected Error: ' + str(err))
            else:
                self._gui.draw_error_message_box(3, 'Success', 'Report generated.')
                self._gui.destroy()

    def _rewind_gui(self):
        self._gui.destroy_generating_label()
        self._gui.show_generate_button()
        self._gui.generate_button.configure(command=self._generate_button_click)

    def _open_button_click(self):
        self._gui.out_directory_entry.delete(0, END)
        self._gui.out_directory_entry.insert(0, os.path.normpath(filedialog.askdirectory()))

    def _validate_input_dir(self) -> bool:
        if os.path.isdir(self._gui.out_directory_entry.get()):
            return True
        else:
            self._gui.draw_error_message_box(2, 'Error', 'Invalid output directory.')
            return False


class ReportGUI:
    # based on code generated from GUI.tcl by PAGE
    def __init__(self, top=None):
        self.root = top
        top.geometry('300x150+700+300')
        top.title('City Report')

        self.intro_label = ttk.Label(top)
        self.intro_label.place(relx=0.13, rely=0.1, height=50, width=200)
        self.intro_label.configure(wraplength='200')
        self.intro_label.configure(text='Enter or select directory to store the file.')

        self.out_directory_label = Label(top)
        self.out_directory_label.place(relx=0.0, rely=0.53, height=21, width=95)
        self.out_directory_label.configure(text='Output Directory')

        self.out_directory_entry = ttk.Entry(top)
        self.out_directory_entry.place(relx=0.35, rely=0.53, relheight=0.14, relwidth=0.42)

        self.open_button = ttk.Button(top)
        self.open_button.place(relx=0.78, rely=0.52, height=24, width=49)
        self.open_button.configure(text='Open...')

        self.generate_button = ttk.Button(top)
        self.generate_button.place(relx=0.37, rely=0.72, height=25, width=76)
        self.generate_button.configure(text='Generate')

        self.generating_label = None

    def destroy(self):
        self.root.destroy()
        self.root = None

    def show_generating_label(self):
        self.generating_label = Label(self.root)
        self.generating_label.place(relx=0.37, rely=0.72, height=21, width=120)
        self.generating_label.configure(text='Report generating...')

    def destroy_generating_label(self):
        self.generating_label.destroy()
        self.generating_label = None

    def show_generate_button(self):
        self.generate_button = ttk.Button(self.root)
        self.generate_button.place(relx=0.37, rely=0.72, height=25, width=76)
        self.generate_button.configure(text='Generate')

    def destroy_generate_button(self):
        self.generate_button.destroy()
        self.generate_button = None

    @staticmethod
    def draw_error_message_box(message_type: int, title: str, message: str):
        if message_type == 1:
            messagebox.showerror(title, message)
        elif message_type == 2:
            messagebox.showwarning(title, message)
        elif message_type == 3:
            messagebox.showinfo(title, message)
        else:
            raise ValueError('Invalid message_type.')


if __name__ == '__main__':
    main()
