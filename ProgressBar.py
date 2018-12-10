__author__ = 'rvoorheis'

import time
import sys

class ProgressBar:

    barLength = 1 # Modify this to change the length of the progress bar
    header = 'Percent:'

    def __init__(self, bl = 10, hdr = 'Percent:'):
        self.barLength = bl
        self.header = hdr

    # update_progress() : Displays or updates a console progress bar
    ## Accepts a float between 0 and 1. Any int will be converted to a float.
    ## A value under 0 represents a 'halt'.
    ## A value at 1 or bigger represents 100%
    def update_progress(self, step, total, currentfile):

        status = ""

        if isinstance(step, int):
            step = float(step)
        if isinstance(total, int):
            total = float(total)

        assert total >= step
        progress = float(step / total)

        if progress < 0:
            progress = 0
            status = "\r\nHalt...\r\n"
        if progress >= 1:
            progress = 1
            status = "\r\nDone...\r\n"
        block = int(round(self.barLength * progress))
        #print ("block = " + str(block) + "  barLength = " + str(self.barLength))

        text = "\r"+ self.header + " [{0}] {1} of {2} {3} {4}".format( "#"*block + "-"*(self.barLength-block), step, total, currentfile, status)
        sys.stdout.write(text)
        sys.stdout.flush()
