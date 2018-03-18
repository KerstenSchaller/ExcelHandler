using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
using System.Reflection;
using System.IO;
using System.Runtime.CompilerServices;
using System.Diagnostics;

namespace ExcelHandler
{
    class ExcelHandler
    {
        Microsoft.Office.Interop.Excel.Application XL;
        Microsoft.Office.Interop.Excel._Workbook WB;
        Microsoft.Office.Interop.Excel._Worksheet Sheet;
        Microsoft.Office.Interop.Excel.Range Rng;

        //current directory
        string curr_dir;

        [MethodImpl(MethodImplOptions.NoInlining)]

        private string GetCurrentMethod()
        {
            StackTrace st = new StackTrace();
            StackFrame sf = st.GetFrame(1);

            return sf.GetMethod().Name;
        }

        private void printError(string method, Exception e)
        {
            Console.WriteLine("-----------------------------------------");
            Console.WriteLine("Error in: " + method  );
            Console.WriteLine("-----------------------------------------");
            Console.WriteLine(e);
            Console.WriteLine("-----------------------------------------");
            Console.WriteLine("-----------------------------------------");
        }

        /*Constructor*/
        public ExcelHandler()
        {   
            // get current directory
            curr_dir = Directory.GetCurrentDirectory();

            try
            {
                // get current active excel instance if present
                XL = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception)
            {
                Console.WriteLine("No excel Instance found-> Opening new");
                //Start Excel and get Application object.
                XL = new Microsoft.Office.Interop.Excel.Application();
            }
            XL.Visible = false;
        }

        public bool OpenFile(string filename)
        {
            try
            {
                Console.WriteLine("Opening File");
                //Get a new workbook.
                //oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));

                //open workbook for reading
                WB = XL.Workbooks.Open(Path.Combine(curr_dir, filename));

            }
            catch (Exception e)
            {

                printError(GetCurrentMethod(), e);
                return false;
            }
            return true;
        }

        public bool OpenWorksheet(string worksheet)
        {
            try
            {
                Console.WriteLine("Opening Worksheet");
                Sheet = (Worksheet)WB.Sheets[worksheet];
            }
            catch (Exception e)
            {
                printError(GetCurrentMethod(), e);
                return false;
            }
            return true;
        }

        public bool WriteToCell(string row, string column, string value)
        {
            try
            {
                Console.WriteLine("Writing Values to Cell");
                Sheet.Cells[row, column] = value;
            }
            catch (Exception e)
            {
                printError(GetCurrentMethod(), e);
                return false;
            }
            return true;
        }

        public bool SavetoFile(string FileName)
        {
            try
            {
                Console.WriteLine("Saving to File");
                string path = Path.Combine(curr_dir, FileName);
                WB.SaveAs(path);
                
            }
            catch (Exception e)
            {
                printError(GetCurrentMethod(), e);
                return false;
            }
            return true;
        }


        public bool CloseFile()
        {
            try
            {
                Console.WriteLine("Closing File");
                WB.Close(0);
                XL.Quit();
            }
            catch (Exception e)
            {

                printError(GetCurrentMethod(), e);
                return false;
            }
            return true;
        }

        


    }


}
