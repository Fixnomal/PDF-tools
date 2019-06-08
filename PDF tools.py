import tkinter
from tkinter import messagebox, filedialog, ttk
from pathlib import Path
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter

def mergeAllPDFs():
    merger = PdfFileMerger()
    for file in Path(str(mergeFolder.get())).iterdir():
        if file.suffix == ".pdf":
            cPdf = open(file, "rb")
            merger.append(cPdf)

    if not str(folderMergePDF.get()).endswith(".pdf"):
        folderMergePDF.set(folderMergePDF.get() + ".pdf")
    with open(Path(str(folderMergePDF.get())),"wb") as folderMergePDFhandle:
        merger.write(folderMergePDFhandle)
    messagebox.showinfo("PDFs of folder merged", f"All PDFs found in {folderMergePDF.get()} merged into {folderMergePDF.get()}")

def merge2PDFs():
    outPutPDF = PdfFileWriter()
    sourcePDF1Handle = PdfFileReader(sourcePDF1.get())
    sourcePDF2Handle = PdfFileReader(sourcePDF2.get())
    numberOfPages1 = sourcePDF1Handle.getNumPages()
    numberOfPages2 = sourcePDF2Handle.getNumPages()
    if numberOfPages1 > numberOfPages2:
        maxNumberOfPages = numberOfPages1
    else:
        maxNumberOfPages = numberOfPages2

    pageIndex = list() #Make a page index that can be regular or reversed and then be used in any case
    if revOrder.get():
        for page in range(numberOfPages2-1,-1,-1):
            pageIndex.append(page)
    else:
        for page in range(numberOfPages2):
            pageIndex.append(page)

    if not interlace.get():
        for pageNum in range(numberOfPages1):
            outPutPDF.addPage(sourcePDF1Handle.getPage(pageNum))
        for pageNum in range(numberOfPages2):
            outPutPDF.addPage(sourcePDF2Handle.getPage(pageIndex[pageNum]))
    if interlace.get():
        for pageNum in range(maxNumberOfPages):
            try: #for different length documents to suppress page number error
                outPutPDF.addPage(sourcePDF1Handle.getPage(pageNum))
            except:
                pass
            try:
                outPutPDF.addPage(sourcePDF2Handle.getPage(pageIndex[pageNum]))
            except:
                pass

    if not mergedPDF.get().endswith(".pdf"):
        mergedPDF.set(mergedPDF.get() + ".pdf")
    with open(mergedPDF.get(),"wb") as pdfWriteHandle:
        outPutPDF.write(pdfWriteHandle)

    messagebox.showinfo("Selected PDFs merged", f"{sourcePDF1.get()} merged with {sourcePDF2.get()}")

def splitPDF():
    outPutPDF = PdfFileWriter()
    splitPDFHandle = PdfFileReader(splitSourcePDF.get())

    if splitPDFRootName.get().endswith(".pdf"):
        splitPDFRootName.set(splitPDFRootName.get()[:-4])

    for pageNum in range(splitPDFHandle.getNumPages()):
        outPutPDF.addPage(splitPDFHandle.getPage(pageNum))
        with open(splitPDFRootName.get() + " page " + str(pageNum + 1) +".pdf", "wb") as pdfWriteHandle:
            outPutPDF.write(pdfWriteHandle)

    messagebox.showinfo("Selected PDFs split", f"{splitSourcePDF.get()} split")

def generalInstructions():
    messagebox.showinfo("Info", "If the programs fails the most common reason is that the PDFs are unreadable for some "
                                "reason unknown to me. A workaround is to print the PDFs using Microsoft PDF writer, "
                                "resulting in an otherwise identical PDF but these are somehow readable. Not ideal I "
                                "know but that's the best advice I got.")

mainWindow = tkinter.Tk()
mainWindow.title("PDF tools")

menuBar = tkinter.Menu(mainWindow)
menuBar.add_command(label="Help", command=generalInstructions)
mainWindow.config(menu=menuBar)

columnsInGrid = 4

# Merge ALL PDFs in folder
segmentStartRow = 0

tkinter.Label(mainWindow, text="Merge all pdf files in the following folder:", font="bold 14").grid(row=0 + segmentStartRow, column=0, columnspan = 2, sticky="w", padx=(0, 20))
mergeFolder = tkinter.StringVar()
mergeFolder.set(Path.home())

tkinter.Label(mainWindow,text="Folder path:", font="bold 10").grid(row=1 + segmentStartRow, column=0, sticky="w")
ttk.Entry(mainWindow, textvariable=mergeFolder).grid(row=1 + segmentStartRow, column=1, sticky="ew")
tkinter.Button(mainWindow, text="Browse", command=lambda: mergeFolder.set(filedialog.askdirectory())).grid(row=1 + segmentStartRow, column=2, sticky="w", padx=5)

folderMergePDF = tkinter.StringVar()
defaultOutputFile = str(Path.home()) + r"\All PDFs in folder merged.pdf"
folderMergePDF.set(defaultOutputFile)
tkinter.Label(mainWindow,text="Select merged output file name:", font="bold 10").grid(row=2 + segmentStartRow, column=0, sticky="w")
ttk.Entry(mainWindow, textvariable=folderMergePDF).grid(row=2 + segmentStartRow, column=1, sticky="ew")
tkinter.Button(mainWindow, text="Browse", command=lambda : folderMergePDF.set(filedialog.asksaveasfilename())).grid(row=2 + segmentStartRow, column=2, sticky="w", padx=5)
tkinter.Button(mainWindow, text="Merge all PDF's in selected folder", command=mergeAllPDFs).grid(row=3 + segmentStartRow, column=1, sticky="ew")
tkinter.Button(mainWindow, text="Cancel", command=quit).grid(row=3 + segmentStartRow, column=0, sticky="e", padx=5)
tkinter.ttk.Separator(mainWindow).grid(row=4 + segmentStartRow, column=0, columnspan=columnsInGrid, sticky="ew", pady=10)

# Merge 2 specific files
segmentStartRow = 5

tkinter.Label(mainWindow, text="Merge the following 2 PDFs:", font="bold 14").grid(row=0 + segmentStartRow, column=0, columnspan = 2, sticky="w", padx=(0, 20))

sourcePDF1 = tkinter.StringVar()
sourcePDF1.set(Path.cwd())
tkinter.Label(mainWindow,text="Select 1st PDF file:", font="bold 10").grid(row=1 + segmentStartRow, column=0, sticky="w")
ttk.Entry(mainWindow, textvariable=sourcePDF1).grid(row=1 + segmentStartRow, column=1, sticky="ew")
tkinter.Button(mainWindow, text="Browse", command=lambda : sourcePDF1.set(filedialog.askopenfilename())).grid(row=1 + segmentStartRow, column=2, sticky="w", padx=5)

sourcePDF2 = tkinter.StringVar()
sourcePDF2.set(Path.cwd())
tkinter.Label(mainWindow,text="Select 2nd PDF file:", font="bold 10").grid(row=2 + segmentStartRow, column=0, sticky="w")
ttk.Entry(mainWindow, textvariable=sourcePDF2).grid(row=2 + segmentStartRow, column=1, sticky="ew")
tkinter.Button(mainWindow, text="Browse", command=lambda : sourcePDF2.set(filedialog.askopenfilename())).grid(row=2 + segmentStartRow, column=2, sticky="w", padx=5)

mergedPDF = tkinter.StringVar()
# currentDirectory = Path()
# currentDirectory = Path.getsourcePDF1.get
defaultOutputFile = r"C:\temp\Test folder\Merge of selected 2 PDFs.pdf" #str(Path.home()) + r"\Merge of selected 2 PDFs.pdf"
mergedPDF.set(defaultOutputFile)
tkinter.Label(mainWindow,text="Select merged output file name:", font="bold 10").grid(row=3 + segmentStartRow, column=0, sticky="w")
ttk.Entry(mainWindow, textvariable=mergedPDF).grid(row=3 + segmentStartRow, column=1, sticky="ew")
tkinter.Button(mainWindow, text="Browse", command=lambda : mergedPDF.set(filedialog.asksaveasfilename())).grid(row=3 + segmentStartRow, column=2, sticky="w", padx=5)
tkinter.Button(mainWindow, text="Merge selected PDFs", command=merge2PDFs).grid(row=4 + segmentStartRow, column=1, sticky="ew")
tkinter.Button(mainWindow, text="Cancel", command=quit).grid(row=4 + segmentStartRow, column=0, sticky="e", padx=5)
tkinter.ttk.Separator(mainWindow).grid(row=5 + segmentStartRow, column=0, columnspan=columnsInGrid, sticky="ew", pady=10)

interlace = tkinter.BooleanVar()
tkinter.Checkbutton(mainWindow, variable=interlace, text="Interlace (Doc1-page1, doc2-page1, doc-page2, etc.)").grid(row=1 +segmentStartRow, column=3, sticky="w")
revOrder = tkinter.BooleanVar()
tkinter.Checkbutton(mainWindow, variable=revOrder, text="reverse page order (2nd doc only)").grid(row=2 +segmentStartRow, column=3, sticky="w")


# Split PDF
segmentStartRow = 11

tkinter.Label(mainWindow, text="Split every page of the following PDF file into individual PDF files:", font="bold 14").grid(row=0 + segmentStartRow, column=0, columnspan=4, sticky="w", padx=(0, 20))
splitSourcePDF = tkinter.StringVar()
splitSourcePDF.set(Path.home())
tkinter.Label(mainWindow,text="Select PDF to split:", font="bold 10").grid(row=1 + segmentStartRow, column=0, sticky="w")
ttk.Entry(mainWindow, textvariable=splitSourcePDF).grid(row=1 + segmentStartRow, column=1, sticky="ew")
tkinter.Button(mainWindow, text="Browse", command=lambda : splitSourcePDF.set(filedialog.askopenfilename())).grid(row=1 + segmentStartRow, column=2, sticky="w", padx=5)

splitPDFRootName = tkinter.StringVar()
defaultOutputFile = str(Path.home()) + r"\Split PDF.pdf"
splitPDFRootName.set(defaultOutputFile)
tkinter.Label(mainWindow,text="Select root name for split fiels:", font="bold 10").grid(row=2 + segmentStartRow, column=0, sticky="w")
ttk.Entry(mainWindow, textvariable=splitPDFRootName).grid(row=2 + segmentStartRow, column=1, sticky="ew")
tkinter.Button(mainWindow, text="Browse", command=lambda: splitPDFRootName.set(filedialog.asksaveasfilename())).grid(row=2 + segmentStartRow, column=2, sticky="w", padx=5)
tkinter.Button(mainWindow, text="Split selected PDF", command=splitPDF).grid(row=3 + segmentStartRow, column=1, sticky="ew")
tkinter.Button(mainWindow, text="Cancel", command=quit).grid(row=3 + segmentStartRow, column=0, sticky="e", padx=5)
tkinter.ttk.Separator(mainWindow).grid(row=4 + segmentStartRow,pady=3)

mainWindow.grid_columnconfigure(1,minsize=300)
mainWindow.mainloop()