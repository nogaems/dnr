# dnr
.doc(x) name restorer

In case of using `testdisk`, `photorec` or any other file recovery tool, often filenames cannot be restored. Usually restored filename looks like "f0738328.docx" or similar. When there are too many files, it could become a real problem to find file you want. To make search a bit easier it is possible to take some text from the file and use it as the name. This could not solve the problem fully, but I think such solution is better than nothing.

## Usage:
```
user@desktop ~/dnr $ ./restorer.py -h
usage: restorer.py [-h] [--input INPUT] [--output OUTPUT] [--verbose]

optional arguments:
  -h, --help            show this help message and exit
  --input INPUT, -i INPUT
                        Path to the directory that contains document files
                        that you was restored
  --output OUTPUT, -o OUTPUT
                        Specified path to save.
  --verbose, -v         Verbose output
```
## Requirements:
* [python-docx] (https://python-docx.readthedocs.io/en/latest/)
* [libreoffice] (https://www.documentfoundation.org/)
* python3

> Beware, you must close open instances of LibreOffice before running this, or it will exit silently without doing anything. This is a known [bug] (https://bugs.documentfoundation.org/show_bug.cgi?id=37531).
