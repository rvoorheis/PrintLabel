__author__ = 'rvoorheis'
# Easy access to excel spreadsheets
import openpyxl

class ExcelControl:

    def __init__(self):
        pass

    def setsheet(self, filename, sheetname):
        """
        Open the input workbook and return the Worksheet object containing the tests
        :param filename: # Input Workbook file name
        :param sheetname: # sheet name in workbook containing tests to run
        :return: Worksheet object to obtain spreadsheet data
        """
        try:
            wb = openpyxl.load_workbook(filename)
            #print (wb.get_sheet_names())
            return wb.get_sheet_by_name(sheetname)

        except Exception as e:
            print ("ExcelControl.setsheet Error " + str(e) + " opening " + filename)
            quit(-8)

    def openworkbook(self, filename):
        """
        Open a workbook
        :param filename: Name of file to open
        :return: wb - Workbook object
        """
        try:
            return openpyxl.load_workbook(filename)

        except ExcelControl as e:
            print ("ExcelControl.openworkbook Error " + str(e) + " opening " + filename)
            quit(-9)
