Excelhandler.exe reads in a template Excel File, opens a worksheet, writes value to it and safes it to an outputfile.

Excelhandler.exe <TemplateFile> <Worksheet><ouputfileName> <Row><Colum><Value>
-> All Inputvalues must be devided by whitespaces, therefore the values are not allowed to contain whitespaces
<TemplateFile>   	should be the complete filename of an excelfile in the same directory as Excelhandler.exe
<Worksheet>		should be the name of a worksheet in the excelfile. 
<ouputfileName>		should be any filename in the same excel file format as the template
<Row><Colum><Value> 	input which specifies a row, column and value to write to it. Multiple values can be specified. 

Example: 
Excelhandler.exe Vorlage.xlsx PersönlicheAusgabenüberwachung Output_2.xlsx 12 F 120 13 F 210 14 F 170




The basic core of the programm is a class called ExcelHandler.cs which is a wrapper for the Microsoft Excel Object Library which makes working with it easier. The class was extended with a console interface to perform above mentioned easy operation.
