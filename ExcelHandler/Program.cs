using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHandler
{
    class Program
    {

        /*
         Excelhandler.exe reads in a template Excel File, opens a worksheet, writes value to it and safes it to an outputfile.

            Excelhandler.exe <TemplateFile> <Worksheet><ouputfileName> <Row><Colum><Value>
            -> All Inputvalues must be devided by whitespaces, therefore the values are not allowed to contain whitespaces
            <TemplateFile>   	should be the complete filename of an excelfile in the same directory as Excelhandler.exe
            <Worksheet>		should be the name of a worksheet in the excelfile. 
            <ouputfileName>		should be any filename in the same excel file format as the template
            <Row><Colum><Value> 	input which specifies a row, column and value to write to it. Multiple values can be specified. 
            
            Example: 
            Excelhandler.exe Vorlage.xlsx PersönlicheAusgabenüberwachung Output_2.xlsx 12 F 120 13 F 210 14 F 170

             */

        static void Main(string[] args)
        {

            if (args == null)
            {
                Console.WriteLine("No arguments given! Try again when you know what you want!");
            }
            else
            {

                ExcelHandler excl = new ExcelHandler();


                excl.OpenFile(args[0]);
                excl.OpenWorksheet(args[1]);

                for (int i = 3; i <= (args.Length - 3); i=i+3 )
                {
                    excl.WriteToCell(args[i], args[i+1], args[i+2]);
                }

                excl.SavetoFile(args[2]);
                excl.CloseFile();
                
            }


            //excl.OpenFile("Vorlage.xlsx");
            //excl.OpenWorksheet("Persönliche Ausgabenüberwachung");
            //excl.WriteToCell("12","F","50");
            //excl.WriteToCell("13", "F", "60");
            //excl.WriteToCell("14", "F", "70");
            //excl.WriteToCell("15", "F", "80");
            //excl.SavetoFile("output.xlsx");
            //excl.CloseFile();




            
        }
    }
}
