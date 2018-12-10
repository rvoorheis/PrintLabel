__author__ = 'rvoorheis'
import os

class FileLocation:
    """
        Navigates the directory structure and provides the fully qualified names for the Archive file
        and the Output files
    """

    outputpath = ""     #Path to Output file location
    archivepath = ""    #Path to Archive file location

    """
    #    Navigates the path to the Archive and Output directories
    #    Archive and output files thus:
    #
    #<common>-|--<Language>-<Application>-<Type>-<Printer>
    #         |    \          \              \    \
    #         |     \          \              \    \-Printer Name
    #         |      \          \              \-Type Archive or Output
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



    def fixextension(self, spath, language):
        work = self.left(spath, len(spath)-3)
        work = work + language
        return work


    def pathend(self, spath, type, prtr, labelname, language):
        """
        Pathend - appends the printer name as the final directory name for files
        :param spath: Static first portion of the path
        :param prtr:     Printer name
        :return: Fully qualified directory name for the Archive and Output files
        """
        try:
            spath = os.path.join(spath, type)   #Output or Archive
            if os.path.exists(spath):           #If directory does not exist,
                pass
            else:
                os.mkdir(spath)                 #Make one

            spath = os.path.join(spath, prtr)
            if os.path.exists(spath):          #If printer directory does not exist,
                pass
            else:
                os.mkdir(spath)                #Make one

            spath = os.path.join(spath, labelname)    # add the label labelname
            spath = self.fixextension(spath, language) # fix the output extension name.

            #print (spath)
            return spath

        except IOError as e:
            print ("Error finding file location " + str(e) + " " + spath)
            quit(-8)

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
#            spath = os.path.dirname(spath)            #Directory containing Label Directory
            spath = os.path.join(spath, rc.Language)           #Down to Language Sub directory
            spath = os.path.join(spath, rc.Application)         #Down to Application Sub directory

            self.outputpath = self.pathend(spath, "Output", rc.Printer, rc.Label, rc.Language)   #add Output branch
            self.archivepath = self.pathend(spath, "Archive", rc.Printer, rc.Label, rc.Language) #Insert Archive Branch

        except Exception:
            print ("Error finding archive file location " + labeldirectory)
            quit(-7)
