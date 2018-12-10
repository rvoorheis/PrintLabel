__author__ = 'rvoorheis'

class RowContents:

    ProcessRow = False
    srow = ""
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
        if Sheet.cell('A'+str(RowNum)).value == "Y":
            self.ProcessRow = True
            self.srow =        str(RowNum)
            self.Printer =     str(Sheet.cell('B'+ self.srow).value)
            self.Dpi =         str(Sheet.cell('C'+ self.srow).value)
            self.Application = str(Sheet.cell('D'+ self.srow).value)
            self.Label =       str(Sheet.cell('E'+ self.srow).value)
            self.Language =    str(Sheet.cell('F'+ self.srow).value)
            #print ('Row ='+self.srow + ' ' + self.Printer +' ' + self.Dpi + ' ' \
            #      + self.Application + ' ' + self.Label + ' ' + self.Language)

        else:
            self.ProcessRow = False