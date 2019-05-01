__author__ = 'rvoorheis'
# -*- encoding: UTF-8 -*-
"""
Requires Python27 and zebra.
"""
#import PrintParameters
import os
from StringIO import StringIO

from zebra.io.unity import SocketConnection
from zebra.proto.parser.grf import parse_grf_to_image

def MakeBMP(self, ZPLfile, BMPfile, IpAddr):
    # coding=utf-8

    import os

    file = open(ZPLfile, 'r')
    Label = file.read()

    # Need to strip off any trailing "^XA^IDR:GED42000.GRF^FS^XZ" strings
    Done = False
    while not Done:
        LabelLength = Label.__len__()
        if LabelLength > 26:

            while Label[LabelLength - 1] == "\n":
                Label = Label[0:LabelLength-1]
                LabelLength = Label.__len__()
            #end loop

            if Label[LabelLength - 26:LabelLength - 20] == "^XA^ID" and Label[LabelLength - 6:LabelLength] == "^FS^XZ":
                Label = Label[0:LabelLength-27]
                LabelLength = Label.__len__()
            else:
                Done = True
            #end if

        else:
            Done = True
        #end if

        Label = Label[0:LabelLength - 3] + ('\n^HCN,,,,,Y^XZ')
#        Label.replace("^XZ", "^HCN,,,,,Y^XZ")


        while createBMP(BMPfile, Label, IpAddr) == -1:
            pass
        #end Loop
        return

def createBMP (BMPfile, LABEL, IpAddr):
    try:
        #printer = SocketConnection("10.80.22.72", 9100)
        printer = SocketConnection(IpAddr, 9100)

        printer.send(LABEL)

        grf_data = printer.collect_sincelast(2.0, 60.0)

        if grf_data == b"":     #Printer problem!!??

            print '********************************************'
            print '  Printer not Available!!  '
            print BMPfile
            print ' Enter "R" to resume or "A" to abort        '
            print '********************************************'
            r = raw_input ("Enter:")
            if r == "r" or "R":
                return (-1)
            else:
                quit(-8)
            #end if
        _, printer_img = parse_grf_to_image(StringIO(grf_data))

        printer_img.save(BMPfile, 'BMP')
        printer.disconnect()
        #print (BMPfile + ' created')

    except Exception as e:
        print ("MakeBMP.createBMP error " + str(e))
        quit(-8)

    return(0)
