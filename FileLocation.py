__author__ = 'rvoorheis'
import os

class FileLocation:
    """
        Navigates the directory structure and provides the fully qualified names for the Archive file
        and the Output files
    """

    outputpath = ""     #Path to Output file location
    archivepath = ""    #Path to Archive file location
    BMPPath = ""        #Path to BMP images

    """
    #    Navigates the path to the Archive and Output directories
    #    Archive and output files thus:
    #
    #<common>-|--<Language>-<Application>-<Type>-<Printer>
    #         |    \          \              \    \
    #         |     \          \              \    \-Printer Name
    #         |      \          \              \-Type Archive or Output or BMP
    #         |       \          \
    #         |        \          \-Name of ZebraDesigner Application (ZD, ZDP, ZDX, ZDS or ZDD)
    #         |         \                                               \   \   \    \      \-Driver with Word or other
    #         |          \-Printer Language (ZPL, EPL or CPL)            \   \   \    \-ZebraDesigner for SAP
    #         |--Labels                                                   \   \   \-ZebraDesigner for XML
    #           <labelname>.lbl                                            \   \-ZebraDesigner Pro
    #                                                                       \-ZebraDesigner
    #
    """

    def left(self, string, count):
        """
        This function is to serve as documentation for the way I implemented the left() function in Python.
        :param string: The string that you want the left most characters of.
        :param count:  Number of characters to return
        :return: The left most (count) characters of the string.
        """
        return string[0:count]


    def __init__(self, labeldirectory, rc):
        """
        Initially set the static portion of the path
        :param labeldirectory: Label file name directory
        :param lang: Printer language
        :param appl: Application name
        :param prtr: Printer name
        :return:
        """
        try:
            spath = os.path.abspath(labeldirectory)   #Absolute path of label directory
#
            spath      = self.AddDir(str(spath), rc.Language)           #Down to Language Sub directory
            spath      = self.AddDir(str(spath), rc.Application)         #Down to Application Sub directory
            spath      = self.AddDir(str(spath), rc.Printer)              #Down to the printer Sub Directory
            self.outputpath = self.AddDir(str(spath), "Output")                #Output Path

            self.archivepath = self.AddDir(str(spath), "Archive")              #Archive Path

            self.BMPPath = self.AddDir(str(spath), "BMP")                  # BMP Path


        except Exception as e:
            print ("Error finding archive file location " + labeldirectory + " " + str(e))
            quit(-7)

    def AddDir (self, FirstPart, SecondPart):

        try:
            spath = os.path.abspath(FirstPart)
            spath = os.path.join (spath, SecondPart)
            if os.path.exists(spath):  # If Archive directory does not exist,
                pass
            else:
                os.mkdir(spath)  # Make one

            return spath

        except Exception as e:
            print ("Error finding Creating directory " + labeldirectory + " "+str(e))
            quit(-9)