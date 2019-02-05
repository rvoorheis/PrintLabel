__author__ = 'rvoorheis'

class RowContents:

    ProcessRow = False
    iRow = 1
    Printer = ""
    Dpi = ""
    Application = ""
    Label = ""
    Language = ""
    OutputPort = ""

    def __init__(self, Sheet, RowNum):
        """

        :rtype : object
        """
        if Sheet.cell(RowNum, 1).value == "Y":
            self.iRow = RowNum
            self.ProcessRow = True

            self.Printer =     str(Sheet.cell(RowNum, 2).value)
            self.Dpi =         str(Sheet.cell(RowNum, 3).value)
            self.Application = str(Sheet.cell(RowNum, 4).value)
            self.Label =       str(Sheet.cell(RowNum, 5).value)
            self.Language =    str(Sheet.cell(RowNum, 6).value)


        else:
            self.ProcessRow = False