excelXML
========

ExcelXML is a java based console application that creates an XML file
from an excel (xlsx) file.  The resulting xml follows a very basic
structure that is the same as if exporting from MS Excel.  

Requirements:
-------------
  * A java virtual machine.
  * Currently testing this with OpenJDK 1.7 under CentOS 6.5.

Source Requirements:
--------------------
Mostly Apache's POI as the API for working with excel files.
Also some maven utilities.
  
  * Non-standard issue java essentials here:
    * poi-3.10-FINAL
    * dom4j-1.6.1
    * Xerces-J-2_11_0
    * xmlbeans-2.6.0

Operation:
----------
  * Download the pre-built excelXML.jar file
  * Alternatively, download the source and compile
  * Run the application supplying two arguments
    * First argument is the name/path of the excel file.
    * Second is the desired name/path of the xml file.
    * Ex. `java -jar excelXML.jar filename.xlsx output.xml`

Reasons for project:
--------------------
  * The current version of MS Excel requires the user to map their worksheet
    to an xml schema or create one.  This is a bit of a chore for some users.
  * There is an old but useful plugin for Excel 2003 but requires minor code
    changes to work.
  * I wanted a way to add this functionality to other apps I am working on.
  