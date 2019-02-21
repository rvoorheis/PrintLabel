from MakeBMP import MakeBMP

__author__ = 'rvoorheis'
# -*- encoding: UTF-8 -*-

import shutil
import os
import ExcelControl
import time
import filecmp
import TempFile
import FileLocation
import subprocess
from datetime import date
import time as _time


class PrintLabel:

    LabelFileDirectory = ""         # Root file directory of label formats
    TempOut = "C:\\Output\\Output.prn"

    def __init__(self, ws, parameters):
        self.LabelFileDirectory = parameters.labelfiledirectory


    def printlabel(self, rc, ws, PtrAddr):
        """
        Print label as specified in the row of the input Workbook
        Completes Output and Archive file path names
        :param rc: Row contents object defining characteristics of the test
        :return:
        """

        try:
            TF = TempFile.TempFile()
            files = FileLocation.FileLocation(self.LabelFileDirectory, rc)

            LabelFileName = os.path.join(self.LabelFileDirectory, rc.Label)
            OutputFileName = files.outputpath + "\\" + files.left(rc.Label, len(rc.Label) - 4) + ".prn"
            ArchiveFileName = files.archivepath + "\\" + files.left(rc.Label, len(rc.Label) - 4) + ".prn"
            BMPFileName = files.BMPPath + "\\" + files.left(rc.Label, len(rc.Label) - 4) + ".bmp"

            if rc.Function == "M":
                if self.FileExists(OutputFileName):
                    MakeBMP(self, OutputFileName, BMPFileName,
                            PtrAddr)  # send the ZPL to a printer to generate the .bmp image
                    self.sheetoutput(rc, ".bmp File Created", ws, BMPFileName, "")
                else:
                    self.sheetoutput(rc, "ZPL File not Found", ws, OutputFileName, "")
            else:
                if self.FileExists(LabelFileName):
                    if self.FileExists(self.TempOut):
                        os.remove(self.TempOut)
                    tfilename = TF.writetempfile(rc, self.TempOut, LabelFileName)     # Use the row contents to build the Temp file
                    zdp = self.LaunchLDA(rc.Application, tfilename)      # print the requested label via ZebraDesigner app

                    if self.WaitForFile(self.TempOut):
                        #shutil.copy(self.TempOut, OutputFileName)
                        shutil.copyfile(self.TempOut, OutputFileName)
                        if self.WaitForFile(OutputFileName):  # Determine test results
                            result = self.CheckOutput(OutputFileName, ArchiveFileName, zdp)
                            self.sheetoutput(rc, result, ws, files.outputpath, files.archivepath)  # Report test results
                            # print(rc.Printer + ' ' + rc.Language + ' ' + rc.Label + ' ' + rc.Dpi + ' ' + result)

                            MakeBMP(self, OutputFileName, BMPFileName,
                                    PtrAddr)  # send the ZPL to a printer to generate the .bmp image
                        else:
                            self.sheetoutput(rc, "ZPL File not copied", ws, LabelFileName, "")
                    else:
                        self.sheetoutput(rc, "ZPL File not created", ws, LabelFileName, "")
                else:
                    self.sheetoutput(rc, "Label File Not Found", ws, LabelFileName,"")

        except Exception as e:
            print ("PrintLabel.printlabel error " + str(e))
            quit(-6)

    def LaunchLDA(self, appName, jobFile):
        """
        Launch the specified ZebraDesigner LDA application with corresponding jobFile
        :param appName: Full path to the ZebraDesigner LDA application to run
        :param jobFile: Job file containing label name and printer name and output port name of label
                        to be created
        :return: Nothing is returned
        """
        try:
            zdp = subprocess.Popen([self.appl_path(appName), os.path.abspath(jobFile),' /SILENT',' /i'])
            return zdp

        except OSError as e:
            print ("PrintLabel.LaunchLDA OS Error - " + str(e))
            quit (-14)

        except ValueError as e:
            print ("PrintLabel.LaunchLDA Value Error - " + str(e))
            quit (-14)

        except Exception as e:
            print ("PrintLabel.LaunchLDA Error - " + str(e))
            quit (-14)

    def path_pre(self):
        if os.environ.get('Program Files(x86)') is None:
            return os.environ.get('Program Files')
        else:
            return os.environ.get('Program Files(x86)')

    def kill_app(self, zdp):
        try:
            if zdp.poll() == "None":
                print("zdp.poll = " + str(zdp.poll()))
                zdp.kill()

        except OSError as e:
            print ("PrintLabel.kill_app OS Error - " + str(e))

        except ValueError as e:
            print ("PrintLabel.kill_app Value Error - " + str(e))

        except Exception as e:
            print ("PrintLabel.kill_app Error - " + str(e))


    def appl_path(self, app_name):
        """
        :param self:
        :param app_name: Name of application to run
        returns: fully qualified name to executable to run
        """
        try:

            spath = os.path.join(os.environ.get("PROGRAMFILES"), "Zebra Technologies")

            if (app_name == "ZebraDesigner"):
                spath = os.path.join(spath, "ZebraDesigner 2", "bin", "Design.exe")
            elif (app_name == "ZebraDesigner Pro"):
                spath = os.path.join(spath , "ZebraDesigner Pro 2", "bin", "Design.exe")
            elif (app_name == "NiceLabelDesigner.exe"):
                spath = os.path.join(os.environ.get("PROGRAMW6432"),"NiceLabel" ,
                                     "NiceLabel 2017" , "bin.net" , "NiceLabelDesigner.exe");

            assert (len(spath)> 15)

            return spath
        except Exception as e:
            print ("PrintLabel.appl_path Error - " + str(e))
            quit (-10)

    def sheetoutput(self, rc, Result, ws, newoutput, archive):
        """
        Create a line of output spreadsheet result detail
        :param rc: input Row contents object
        :param Result: String result of test
        :param  ws - Worksheet object of spread sheet
        :return: nothing returned
        """
        try:
            ws.cell(rc.iRow, 7).value = date.today().strftime('%x')
            ws.cell(rc.iRow, 8).value = time.strftime('%X')
            ws.cell(rc.iRow, 9).value = Result
            ws.cell(rc.iRow, 10).value = newoutput
            if Result == "New":
                ws.cell(rc.iRow, 10).value = ''
                ws.cell(rc.iRow, 11).value = archive
            elif Result == "Fail":
                ws.cell(rc.iRow, 10).value = newoutput
                ws.cell(rc.iRow, 11).value = archive
            else:
                ws.cell(rc.iRow, 11).value = ''

        except Exception as e:
            print ("PrintLabel.sheetoutput Error - " + str(e))
            quit (-11)


    def FileExists(self, filename):
        # type: (object, object) -> object
        """
        Checks to see if file exists.
        :param filename: Fully qualified file name of to check for
        :return: result:  Logical True or False whether or not the specified file exists
        """
        try:
            open(filename, 'r')      #open the file
            return (True)

        except IOError as e:
            #print ("PrintLabel.FileExists - Error - " + str(e))
            return (False)          #IO error = assume file does not exist

    def WaitForFile(self, filename):
        """
        Wait for file to exist.  The Application may require time to create the output.
        :param filename: Fully qualified pathname of file to be created.
        :return: reult: Logical True or False, file was created or was not in 10 seconds
        """
        try:
            #print ("waiting for " + filename + " to exist")
            result = True
            for i in range(1,10):
                #print i
                time.sleep(1)
                if self.FileExists(filename):
                    #print(filename + " Found")
                    result = True
                    break
                else:
                    result = False
                    #print (filename + " Never found")

        except IOError as e:
            print "IO Error" + str(e)+ " "+os.strerror
            result = False
            #quit (-9)

        finally:
            return (result)


    def CheckOutput(self, outputfn, archivefn, zdp):
        """
        See if the newly generated output matches the "good" output archive.
        :param outputfn: - Newly generated output file
        :param archivefn: - Full file name of Archive file
        :return: result   - 4 Character description of result of test, suitable for inserting in
                            output spreadsheet or error log
        """
        try:

            result = ""
            if self.WaitForFile(outputfn):                  #Wait for output file to be created
                if self.FileExists(archivefn):
                    if (filecmp.cmp(outputfn, archivefn)):  #Pass -Output matches Archive
                        result = "Pass"
                    else:
                        result = "Fail"                     #Fail - Output does not match
                else:                           #No Archive file exists
                    result = "New"
                    shutil.copy(outputfn, archivefn)
            else:                                               #Output file never created
                result = "Label Err"
                self.kill_app(zdp)

        except Exception as e:
            print ("IO Error" + str(e)+ " "+os.strerror)
            result = "IO Err"

        finally:
            return result


