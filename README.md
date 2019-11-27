# abcl\_excel\_gen
A small program to generate an excel file from a metafile.

## Setup the environment.
The base requirements:

   * a Java runtime
   * Armed Bear Common Lisp (tested with version 1.6.0)
   * Apache Poi (tested with version 4.1.1)

As a convenience, the _setupenv.sh_ script will download and extract the libraries automatically.

## Create an Excel sheet
To run the example, you'll need a command similar to the following:
```java -jar abcl-bin-1.6.0/abcl.jar --noinform -- ~/ss/poi-4.1.1/poi-4.1.1.jar xyz.xlsx xyz.lisp < buildsheet.lisp```

or just look at the Makefile or run `make` to do it automatically.
