import tkinter
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import *
from tkinter import messagebox
from tkinter.ttk import *
from tkinter import filedialog


def show_exception_msg(message):
    messagebox.showerror("Exception", message)


def main():
    window = Tk()
    window.title("Histogram generator")
    window.geometry("505x400")
    binNumber = tkinter.StringVar()
    histogramTitle = tkinter.StringVar()

    def selectFile():
        filetypes = (
            ('Excel files', '*.xlsx'),
            ('Old excel files', '*.xls'),
            ('All files', '*.*')
        )

        filename = filedialog.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)
        try:
            pd.read_excel(filename)
            print("File read successfully")
        except FileNotFoundError:
            show_exception_msg("File Not Found")
            return
        except PermissionError:
            show_exception_msg("Can't access this file, permission error")
            return
        except Exception as erMsg:
            show_exception_msg("An error occurred:", erMsg)
            return

        def createHistogram():
            data = pd.read_excel(filename)
            try:
                if int(binNumber.get()) < 0:
                    show_exception_msg("Input Error, please enter an integer number")
                    return
            except ValueError:
                show_exception_msg("Input Error, please enter an integer number")
                return
            dataArr = []
            labelArr = []
            for col in range(1, len(data.columns.values)):
                dataArr.append(data[data.columns.values[col]].to_numpy())
                labelArr.append(data.columns.values[col])

            dataMin = []
            dataMax = []
            for data in dataArr:
                dataMin.append(min(data))
                dataMax.append(max(data))

            xStart = min(dataMin)
            xStop = max(dataMax) + 1
            xStep = (max(dataMax) - min(dataMin)) / int(binNumber.get())
            plt.xticks(np.arange(start=xStart, stop=xStop, step=xStep))
            try:
                plt.hist(dataArr, int(binNumber.get()), histtype='bar', label=labelArr, density=False,
                         edgecolor='black')
                plt.legend(loc='upper right')
                plt.title(str(histogramTitle.get()))
                plt.show()
            except Exception as e:
                show_exception_msg("Incorrect data format in file, your 1st column or 1st row is empty")

        createButton = Button(
            window,
            text="Create a histogram",
            command=createHistogram
        )

        createButton.grid(row=4, column=0, columnspan=3, sticky=NS)

    title = Label(
        window,
        text="Histogram creator",
        font=("Verdana", 20, 'bold')
    )

    fileLabel = Label(
        window,
        text="Choose an Excel file to create a histogram from:")

    fileLabel.config(font=("Verdana", 12))

    openButton = Button(
        window,
        text='Open a File',
        command=selectFile
    )

    histTitleLabel = Label(
        window,
        text="Name your histogram: ",
        font=("Verdana", 12)
    )

    histTextBox = Entry(
        window,
        font=("Verdana", 12),
        textvariable=histogramTitle,
        width=20
    )

    binNumLabel = Label(
        window,
        text="Input number of bins: ",
        font=("Verdana", 12)
    )

    binNumInput = Entry(
        window,
        width=5,
        textvariable=binNumber,
        font=("Verdana", 12)
    )

    binNumber.set("5")
    histogramTitle.set("Histogram")

    title.grid(row=0, column=0, columnspan=3, sticky=NS, pady=12, padx=5)
    fileLabel.grid(row=1, column=0, columnspan=2, pady=3, sticky=W, padx=5)
    openButton.grid(row=1, column=2, sticky=NS, pady=3, padx=5)
    histTitleLabel.grid(row=2, column=0, sticky=W, pady=3, padx=5)
    histTextBox.grid(row=2, column=1, columnspan=2, sticky=E, pady=3, padx=5)
    binNumLabel.grid(row=3, column=0, columnspan=2, sticky=W, pady=3, padx=5)
    binNumInput.grid(row=3, column=2, sticky=E, pady=3, padx=5)

    window.mainloop()


if __name__ == "__main__":
    main()
