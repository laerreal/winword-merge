#!/usr/bin/python3

DEBUG = 1

WINEPREFIX = None
unix_vbs = "diff.vbs"

from sys import argv, stderr, stdout
from subprocess import Popen, PIPE
from os import environ

if DEBUG < 1:
    log = open("log.txt", "w")

def el(l):
    if DEBUG < 1:
        log.write(l + "\n")
        log.flush()
    else:
        stderr.write(l + "\n")

def pl(*l):
    l = str(l)
    if DEBUG < 1:
        log.write(l + "\n")
        log.flush()
    else:
        stdout.write(l + "\n") 

def winpath(unixpath):
    p = Popen(["winepath", "--windows", unixpath], stdout=PIPE)
    ret, _ = p.communicate()
    if p.returncode != 0:
        el("winepath returned %d" % p.returncode)
        exit(1)
    # strip \n
    return ret[:-1].decode("utf-8")

if __name__ == "__main__":
    argc = len(argv)
    if argc < 2:
        el("Too few argumetns")
        exit(1)
    elif argc == 2:
        mode = "git-resolve"
    elif argc == 3:
        mode = "compare-and-merge"
    else:
        el("Too many arguments %d" % argc)
        el("argv: %s" % argv)
        exit(1)

    f_unix = argv[1]

    if mode == "git-resolve":
        f0_unix = "ours." + f_unix

        ours_cmd = ["git", "show", ":2:" + f_unix]
        if DEBUG < 1:
            pl(ours_cmd)

        ours_p = Popen(ours_cmd, stdout = PIPE, stderr = PIPE)
        ours, ours_err = ours_p.communicate()

        if ours_err:
            el(ours_err.decode("utf-8"))

        if ours_p.returncode:
            el("Failed to get ours version")
            exit(1)

        # XXX: this is wrong way for big files
        ours_f = open(f0_unix, "wb")
        ours_f.write(ours)
        ours_f.close()
    elif mode == "compare-and-merge":
        f0_unix = f_unix
    else:
        el("Unknown mode %s" % mode)
        exit(1)

    if mode == "git-resolve":
        f1_unix = "theirs." + f_unix

        theirs_cmd = ["git", "show", ":3:" + f_unix]
        if DEBUG < 1:
            pl(theirs_cmd)

        theirs_p = Popen(theirs_cmd, stdout = PIPE, stderr = PIPE)
        theirs, theirs_err = theirs_p.communicate()

        if theirs_err:
            el(theirs_err.decode("utf-8"))

        if theirs_p.returncode:
            el("Failed to get theirs version")
            exit(1)

        # XXX: this is wrong way for big files
        theirs_f = open(f1_unix, "wb")
        theirs_f.write(theirs)
        theirs_f.close()
    elif mode == "compare-and-merge":
        f1_unix = argv[2]
    else:
        el("Unknown mode %s" % mode)
        exit(1)

    f0 = winpath(f0_unix)
    f1 = winpath(f1_unix)

    if DEBUG < 1:
        pl(f0,f1)

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

objWord.ActiveDocument.{operation} f1

objWord.ActiveDocument.Close

End Sub
"""

    code = code_template.format(
        f0 = f0,
        f1 = f1,
        operation = {
            "compare-and-merge" : "Compare",
            "git-resolve" : "Merge"
        }[mode]
    )

    vbs = open(unix_vbs, "w")
    vbs.write(code)
    vbs.close()

    # https://stackoverflow.com/questions/2231227/python-subprocess-popen-with-a-modified-environment
    e = environ.copy()
    if WINEPREFIX is not None:
        e["WINEPREFIX"] = WINEPREFIX

    win_vbs = winpath(unix_vbs)

    # https://leereid.wordpress.com/2011/08/03/how-to-run-vbscript/
    """
    ### XXX: fixme WINEPREFIX
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

