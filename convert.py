import PyPDF2
import sys
import re
import csv
import os


class HeightedRow:
    def __init__(self, pageNum, y, data):
        self.pageNum = pageNum
        self.y = y
        self.data = data

    def sort_key(self):
        return (self.pageNum, -self.y)

    def writeToCSV(self, writer):
        writer.writerow(self.data)

    def __repr__(self):
        return "pageNum: " + str(self.pageNum) + " y: " + str(self.y) + " data: " + repr(self.data)


class Entry:
    def __init__(self, loc, text):
        self.loc = loc
        self.text = text

    def __repr__(self):
        return repr((self.loc, self.text))


def getLocTextMap(page):
    content = page.getContents()
    if not isinstance(content, PyPDF2.pdf.ContentStream):
        content = PyPDF2.pdf.ContentStream(content, page.pdf)

    M = {}
    loc = (0,0)
    for operands, operator in content.operations:
        if operator == "Tm":
            loc = (float(operands[4]), float(operands[5]))
        if operator == "TJ":
            M[loc] = operands[0][0]

    return M


def getLocTextMapList(filepath):
    pages = []

    with open(filepath, 'rb') as pdfFileObj:
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        for pageNum in range(0, pdfReader.getNumPages()):
            pages.append(getLocTextMap(pdfReader.getPage(pageNum)))

    return pages


def sortHeightedRows(heightedRows):
    heightedRows.sort(key=lambda x:x.sort_key())


def getHeightedRows(locTextMapList):
    heightedRows = []

    for pageNum, M in enumerate(locTextMapList):
        rows = {}
        for loc in M:
            text = M[loc]
            x = loc[0]
            if re.match(r"\d+/\d+$", text) != None and x > 70 and x < 71:
                rows[loc[1]] = []

        for loc in M:
            text = M[loc]
            if rows.has_key(loc[1]):
                rows[loc[1]].append(Entry(loc, text))

        for r in rows:
            date = ""
            description = ""
            checkNo = ""
            deposits = ""
            withdrawls = ""
            balance = ""

            for c in rows[r]:
                x = c.loc[0]

                if x > 70 and x < 71:
                    date = c.text

                if x > 90 and x < 120:
                    description = c.text

                if x > 350 and x < 370:
                    checkNo = c.text

                if x > 390 and x < 430:
                    deposits = c.text

                if x > 450 and x < 500:
                    withdrawls = c.text

                if x > 510 and x < 600:
                    balance = c.text

            heightedRows.append(HeightedRow(pageNum, r,
                [date, description, checkNo, deposits, withdrawls, balance]))

    return heightedRows


def getAccountMarkers(locTextMapList):
    heightedRows = []

    for pageNum, M in enumerate(locTextMapList):
        rows = {}
        for loc in M:
            text = M[loc]
            x = loc[0]
            if re.match(r"Account number:", text) != None:
                rows[loc[1]] = []

        for loc in M:
            text = M[loc]
            if rows.has_key(loc[1]):
                rows[loc[1]].append(text)

        for r in rows:
            heightedRows.append(HeightedRow(pageNum, r, list(reversed(sorted(rows[r])))))

    return heightedRows


def writeHeightedRowsToCSV(heightedRows, outputpath):
    csvfile = open(outputpath, "w")
    writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

    writer.writerow([
        "date",
        "description",
        "checkNo",
        "deposits",
        "withdrawls",
        "balance",])

    for hr in heightedRows:
        hr.writeToCSV(writer)

    csvfile.close()


if sys.argv[1].endswith('.pdf'):
    filepath = sys.argv[1]

    filedir, filename = os.path.split(filepath)

    outputpath = os.path.join(filedir, filename.replace(".pdf", ".csv"))

    LocTextMapList = getLocTextMapList(filepath)

    heightedRows = getHeightedRows(LocTextMapList) + getAccountMarkers(LocTextMapList)
    sortHeightedRows(heightedRows)

    writeHeightedRowsToCSV(heightedRows, outputpath)


else:
    print("Wrong number of arguments.  Give me a PDF file ending in .pdf")

