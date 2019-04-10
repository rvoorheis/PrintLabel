__author__ = 'rvoorheis'
# -*- encoding: UTF-8 -*-
import PrintParameters
import PrintLabel
import ExcelControl
import RowContents
import ProgressBar
import socket
import MakeBMP
import sys

def main(argv):

    global debug                    #debug flag

    debug = False

    try:

        parms = PrintParameters.PrintParameters(argv)
        if (is_valid_ipv4_address(parms.PrinterAddress)):
            PrtAddr = parms.PrinterAddress
        else:
            print ("RunTests: Invalid IP Address provided. " + parms.PrinterAddress)
            quit(-9)

        parms.Display_parameters()

        w = ExcelControl.ExcelControl()
        wb = w.openworkbook(parms.wbInputFileName)
        ws = wb.get_sheet_by_name("Sheet1")

        pl = PrintLabel.PrintLabel(ws, parms)
        pb = ProgressBar.ProgressBar(25, "Progress:")

        ToDo = 0
        RowCounter = 0
        LastRow = ws.max_row + 1

        for row in range(2, LastRow):
            rc = RowContents.RowContents(ws, row)
            if rc.ProcessRow:
                ToDo = ToDo + 1

        print("reading rows")
        for row in range(2, LastRow):

            assert row <= LastRow

            rc = RowContents.RowContents(ws, row)
            if rc.ProcessRow:
                # print (str(row-1) + " of " + str(counter))
                pl.printlabel(rc, ws, PrtAddr)
                RowCounter = RowCounter + 1
                pb.update_progress(float(RowCounter), float(ToDo), str(rc.Label))
                wb.save(parms.wbInputFileName)

    except LookupError as e:
        print("Lookup Error " + str(e))
        quit(-2)


    except NameError as e:
        print("Error! - Invalid Filename = " + str(e))
        quit(-3)

    except StandardError as e:
        print("Error! - Standard Error = " + str(e))
        quit(-3)

    except Warning as e:
        print("Warning " + str(e))
        quit(-1)

def is_valid_ipv4_address(address):
    try:
        socket.inet_pton(socket.AF_INET, address)
    except AttributeError:  # no inet_pton here, sorry
        try:
            socket.inet_aton(address)
        except socket.error:
            return False
        return address.count('.') == 3
    except socket.error:  # not a valid address
        return False

    return True

if __name__ == '__main__':
    main(sys.argv[1:])