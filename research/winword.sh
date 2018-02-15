#!/bin/bash

f=/media/data/Docs/2.doc

export WINEPREFIX="/home/real/wine32"

wine \
    C:\\\\windows\\\\command\\\\start.exe \
    /Unix /home/real/wine32/dosdevices/c:/Program\ Files/Microsoft\ Office/Office12/winword.exe \
    $(winepath --windows $f) \

