__author__ = 'rvoorheis'

import os

class TempFile:
    tempfilename = ""

    def __init__(self):
        self.tempfilename = "temp.job"

    def writetempfile(self, rc, portname, labelfilename):

        try:

            f = open(self.tempfilename, mode="w")

            f.write('LABEL "'   + labelfilename + '", "' + rc.Printer + '"\n')
            f.write('PRINTER "' + rc.Printer + '"\n')
            f.write('PORT "'    + portname   + '"\n')
            f.write('PRINT 1\n')
            f.write('QUIT\n')

            f.close()
            return self.tempfilename

        except Exception as e:
            print ("Error Creating TempFile " + self.tempfilename)
            print str(e)
            quit(-4)