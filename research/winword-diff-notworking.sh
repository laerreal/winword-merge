#!/bin/bash

f1=/media/data/Docs/1.doc
f0=/media/data/Docs/2.doc
macros=/home/real/me/winword/diffmacros.dotm

export WINEPREFIX="/home/real/wine32"

wine \
    C:\\\\windows\\\\command\\\\start.exe \
    /Unix /home/real/wine32/dosdevices/c:/Program\ Files/Microsoft\ Office/Office12/winword.exe \
    $(winepath --windows $macros) \
    $(winepath --windows $f0) \
    $(winepath --windows $f1) \
    /mCompareAndMerge



