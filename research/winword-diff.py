#!/usr/bin/python3

DEBUG = 0

WINEPREFIX = "/home/real/wine32"
unix_vbs = "diff.vbs"

from sys import argv, stderr
from subprocess import Popen, PIPE
from os import environ

def el(l):
    stderr.write(l + "\n")

def winpath(unixpath):
    p = Popen(["winepath", "--windows", unixpath], stdout=PIPE)
    ret, _ = p.communicate()
    if p.returncode != 0:
        el("winepath returned %d" % p.returncode)
        exit(1)
    # strip \n
    return ret[:-1].decode("utf-8")

if __name__ == "__main__":
    if len(argv) < 2:
        el("Too few argumetns")
        exit(1)

    f0 = winpath(argv[1])
    f1 = winpath(argv[2])

    if DEBUG < 1:
        print(f0,f1)

    code_template = """
Option Explicit

On Error Resume Next

Diff

Sub Diff

Dim objWord
Dim objDoc
Dim f0
Dim f1

f0 = "{f0}"
f1 = "{f1}"

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

objWord.Documents.Open(f0)

objWord.ActiveDocument.Merge f1

objWord.ActiveDocument.Close

End Sub
"""

    code = code_template.format(
        f0 = f0,
        f1 = f1,
    )

    vbs = open(unix_vbs, "w")
    vbs.write(code)
    vbs.close()

    # https://stackoverflow.com/questions/2231227/python-subprocess-popen-with-a-modified-environment
    e = environ.copy()
    e["WINEPREFIX"] = WINEPREFIX

    win_vbs = winpath(unix_vbs)

    # https://leereid.wordpress.com/2011/08/03/how-to-run-vbscript/
    """
    cscript = Popen(
        [
            "wine",
            "C:\\windows\\command\\start.exe",
            "/Unix",
            WINEPREFIX + "/dosdevices/c:/Windows/system32/cscript",
            win_vbs,
        ],
        env = e
    )
    """

    # https://gist.github.com/Ravenstine/9ccbe8269e8d5d28305b

    cscript = Popen(
        [
            "wine",
            "wscript",
            unix_vbs,
        ],
        env = e
    )


    cscript.communicate()

