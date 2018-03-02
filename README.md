# MS Office Word Merging Helper

The script `winword-merge.py` automates merging MS Word documents which
are managed by Git. It is designed to be used as a Nautilus (the Ubuntu Linux
Unity file explorer) script
(`~/.local/share/nautilus/scripts`). The script is to be applied (through
popup menu) to a document which is
_being_ merged. I.e. a `git merge` command was executed and a conflict
with the document is issued. The script launches MS Word document merging UI
using a self-generated Visual Basic script (macros). A user must have Wine
(Windows environment emulator) with MS Word installed. Also, the user
definitely must check `WINEPREFIX` _variable on the top of the script_ (
google what is WINEPREFIX).

This script may be used standalone (without Nautilus).

```bash
cd /path/to/directory/managed/by/git
winword-merge.py conflicting/document.doc
```

## Comparison

Given 2 documents the script launches MS Word document comparison.

### Nautilus

Select two files and use the script.
Note that the order is significant and may be affected by selection of file
sorting criteria (Name, Size, etc.).

> File selection order is seems to do not affect the order of arguments passed
> to the script.

### Standalone

```bash
cd /path/to/a/directory/
winword-merge.py old.doc new.doc
```

## Debugging

There is `DEBUG` variable in the script. Less value results in more prints.

## Testing

The script was tested in next environment:

* Ubuntu Linux 16.04 64-bit
* wine-1.6.2
* 32-bit wine prefix
* MS Word 2007

