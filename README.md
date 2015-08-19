# NIEM-XSD-to-Excel-Converter

Version v2.0 by Thomas Rosales

Build using Windows Forms in C#. This tool is used to generate Excel Spreadsheets from NIEM (XSD) schemas. Compactible with extension, exchange, subset and codelist schemas.

It currently organized the schema contents into excel columns :

- Class Name (Extension Class)
- Element Name	
- Element Type
- Documentation
- Source

#####Instructions:
- Run exe (no installation required)
- Press "Select NIEM XSD" button and choose NIEM schema file 
- Press "Convert To Excel" button.
- Wait for operation to finish (should be quick)
- Output file should appear in same directory as exe tool. 
- The tool outfile will be named "BasicTable.xlsx".

#####Latest changes:

- Embeded DLLs for portability.
- Faster processing. No longer requires Excel.
- Better error catching.

