# NIEM-XSD-to-Excel-Converter

Lastest version v2.2

as of 8/20/2015 

Author: Thomas Rosales

Build using Windows Forms in C#. This tool is used to generate Excel Spreadsheets from NIEM (XSD) schemas. Compactible with extension, exchange, subset and codelist schemas.

It currently organized the schema contents into excel columns :

- Class Name (Extension Class)
- Element Name	
- Element Type
- Documentation***
- Source(opcional)***


#####***Schema Design:
-In order for the tool to capture the Documentation and Source correctly, an element should have a separate documentation tag for Documentation and another one for Source. **Important!**: The source documentation must start with **"Source:"** 

Example:

```
  <xsd:element abstract="false" name="ProgramType" nillable="false" type="niem-xsd:string">
    <xsd:annotation>
      <xsd:documentation>Source: North Dakota</xsd:documentation>
      <xsd:documentation>Referral Type	Indicator of what system is sending this referral to STARS</xsd:documentation>
    </xsd:annotation>
  </xsd:element>
```

#####Instructions:
- Run exe (no installation required)
- Press "Select NIEM XSD" button and choose NIEM schema file 
- Press "Convert To Excel" button.
- Wait for operation to finish (should be quick)
- Once done the spreadsheet will appear in the same directory as the schema selected.


#####Disclamer:
- This tool should be taken lightly. Once the spreadsheet is generated, it needs to be reviewed manually to make sure all elements were included. The tool does its best to capture everything inside the schema. 


#####Download Link:
- https://github.com/Thwar/NIEM-XSD-to-Excel-Converter/blob/master/NIEMXML/App/NIEMXML.exe?raw=true

#####Changes:
#####v2.2
- Fixed no sequence errors
- Modified file output process and message box
- Tool now captures name and ref element attributes
- Fixed cosmetic spreadsheet background color
- Added schema definition
- Source is now allowed without having description field.
- Class now supports source definitions. 


#####v2.0
- Embeded DLLs for portability.
- Faster processing. No longer requires Excel.
- Better error catching.
- Spreadsheet redesign
- Added Source column

