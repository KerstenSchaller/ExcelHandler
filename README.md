# Excelhandler  

The **Excelhandler** reads in a template Excel File, opens a worksheet, writes values to it and saves it to an output file.  

The basic core of the programm is a class called ExcelHandler.cs which is a wrapper for the Microsoft Excel Object Library which makes working with it easier. The class was extended with a console interface to perform above mentioned easy operation.  

## How to use **Excelhandler**  

```
Excelhandler.exe [templateFile] [worksheet] [ouputfileName] [row] [colum] [value]  
```

### Attention  

All input values must be divided by whitespaces, therefore the values are not allowed to contain whitespaces.  

### Parameter  

| Parameter | Description |  
|:-------------:|:-------------|  
| TemplateFile  |Should be the complete filename of an excelfile in the same directory as Excelhandler.exe.|  
|   Worksheet   |Should be the name of a worksheet in the excelfile.|  
| ouputfileName |Should be any filename in the same excel file format as the template.|  
|row colum value|Input which specifies a row, column and value to write to it. Multiple values can be specified.|  

## Example  

Excelhandler.exe Vorlage.xlsx PersönlicheAusgabenüberwachung Output_2.xlsx 12 F 120 13 F 210 14 F 170  
