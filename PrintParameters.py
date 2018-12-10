__author__ = 'rvoorheis'

import os
import sys
import getopt

class PrintParameters:
    """
    Obtain runtime parameters
    """
#    wbInputFileName = "E:\\PythonTestScripts\\input.xlsx"   #Name of input spreadsheet
#    outputport = "C:\\Output\\Output.prn"             #file name port that printer will print to
#    labelfiledirectory = "E:\\ATF\\Labels"          #Directory name containing Label format

    wbInputFileName = ""  # Name of input spreadsheet
    outputport = ""  # file name port that printer will print to
    labelfiledirectory = ""  # Directory name containing Label format

    def __init__(self, argv):

        try:
            opts, args = getopt.getopt(sys.argv[1:], 'hi:l:p:d', ['help', 'input=', 'labels=', 'printout=','debug'])

        except getopt.GetoptError as e:
            print("GetoptError " + str(e))
            self.usage()
            sys.exit(2)
        try:
            if len(opts) == 0:
                self.usage()
                sys.exit(2)
            else:
                for opt, arg in opts:
                    if opt in ("-h", "--help"):
                        self.usage()
                        sys.exit()
                    elif opt in ('-d','--debug'):
                        _debug = True
                    elif opt in ("-i", "--input"):
                        self.wbInputFileName = arg
                    elif opt in ("-p", "--printout"):
                        self.outputport = arg
                    elif opt in ("-l", "--labels"):
                        self.labelfiledirectory = arg

                self.Display_parameters

        except LookupError as e:
            print("Lookup Error " + str(e))
            quit(-2)

        except StandardError as e:
            print("Error! - Standard Error = " + str(e))
            quit(-3)

        except Warning as e:
            print("Warning " + str(e))
            quit(-1)

    def Display_parameters(self):
        print ("Input spreadsheet    = " + self.wbInputFileName)
        print ("Output port          = " + self.outputport )
        print ("Label file directory = " + self.labelfiledirectory)


    def ArchiveLocation(LabelDirectory, Lang, Appl, Prtr ):
        try:
            sPath = os.path.abspath(LabelDirectory)
            sPath = os.path.dirname(sPath)
            sPath = sPath.join(sPath, Lang)
            sPath = os.path.join(sPath, Appl)
            sPath = os.path.join(sPath, "Archive")
            if os.path.exists(sPath):
                pass
            else:
                os.mkdir(sPath)
            sPath = os.path.join(sPath, Prtr)
            if os.path.exists(sPath):
                pass
            else:
                os.mkdir(sPath)
            #print (sPath)
            return sPath

        except Exception as e:
            print ("PrintParameters.ArchiveLocation - Path Error" + str(e))
            quit (-3)

    def usage(self):
            print "PrintLabel"
            print ""
            print "PrintLabel is an application to print ZebraDesigner label format files and compare the resulting output "
            print " against an Archive of 'good' output.  It is driven by a spreadsheet that identifies the label, application"
            print " and printer to use.  It requires a specific directory structure of Label formats, Archive files and Output files"
            print ""
            print "Usage:  PrintLabel --input=<input spreadsheet workbook>"
            print "                   --PrintOut=<File defined as Local port for printers>"
            print "                   --labels=<Path to root of directory structure containing labels and Archive files>"
            print " example:"
            print " Python RunTests.py --input=""E:\PythonTestScripts\input.xlsx --PrintOut=c:\Output\Output.txt  --labels=e:\Atf"
            print ""
            print "The input spreadsheet workbook must have the following parameter columns defined on sheet1"
            print " |Active	 |Printer	     | dpi |  Application      | Label        | Language           |"
            print " |--------|---------------|-----|-------------------|--------------|--------------------|"
            print " |<Y or N>|<Printer Name> |<dpi>|<Application Name> | <Label name> | <Printer Language> |"
            print " "
            print " Each row defines a label a label design application and printer"
            print " The Active column contains a Y or N flag indicating whether to process that row"
            print " The Printer column contains the name of the installed ZebraDesigner printer"
            print "     This printer should have the output directed to a local port and bidirectional input turned off"
            print "     For ZPL printers, the parameters should be set to 'Use Printer Settings'"
            print "  The dpi column specifies the print density of the printer.  This is informational and does not affect processing"
            print "  The Application column specifies the ZebraDesigner application to use to print the label format"
            print "  The Label column identifies the label format to be printed"
            print "  The Language column identifies the printer language.  It should match the printer and used to organize the output"
            print ""
            print "  After execution, the columns to the right of the input column are used to record the results"
            print "  These columns are as follows Date of run; time of run; run results (Pass, Fail, Label error or Label not found)"
            print "      if the result is Fail, the next two columns contain the path of the generated output files and the Archive file "
            print "      That it was compared to. If the result is Label Not Found, the column contains the path to the label format"
            print " "
            print " The local port file must be set to the same file for all printers in the test run."
            print ""
            print " The labels directory structure is laid out as follows"
            print ""
            print ""
            print "   /ATF          Top of structure - This is the directory to identify in the '-Labels=' parameter"
            print ""
            print "			/Labels									<- Holds the test labels."
            print "					/ZPL	<- ZPL Labels"
            print "					/CPCL   <- CPCL Labels"
            print "					/EPL	<- ZPL Labels"
            print "	"
            print "			/ZPL	/ZebraDesigner		/Archive	/Printer Names		The Archive subdirectory leads to where the good label"
            print "																		printer commands for each printer are stored for each label."
            print "										/Output		/Printer Names		The Output subdirectory leads to where printer code is stored"
            print "																		for labels that fail to match the archive code when printed."
            print "					/ZebraDesigner Pro	/Archive	/Printer Names"
            print "										/Output		/Printer Names"
            print "					"
            print "			/EPL	/ZebraDesigner		/Archive	/Printer Names"
            print "										/Output		/Printer Names"
            print "					/ZebraDesigner Pro	/Archive	/Printer Names"
            print "										/Output		/Printer Names"
            print "					"
            print "			/CPCL	/ZebraDesigner		/Archive	/Printer Names"
            print "										/Output		/Printer Names"
            print "					/ZebraDesigner Pro	/Archive	/Printer Names"
            print "										/Output		/Printer Names"
            print ""
            print "			/LabelWork	<- Repository of labels with spreadsheet that describes each label and what they are intended to test."
            print "						<- This directory is not used for Print Label execution but is intended as a Label"
            print "							workarea.  The actual labels used in the tests are in the /Labels/<Printer Language>"
            print "							directories"
            print "							"
            print ""
            print ""


