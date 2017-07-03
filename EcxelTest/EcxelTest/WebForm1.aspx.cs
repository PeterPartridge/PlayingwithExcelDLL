using Microsoft.Office.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;


namespace EcxelTest
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            play();
        }
        public void play()
        {
        //create a list of the object MyItems
            List<MyItems> items = new List<MyItems>();
            //open an instance of the excel application
            Excel.Application exclApp = new Microsoft.Office.Interop.Excel.Application();
            // open the spreadsheet
            Excel.Workbook wkbook = exclApp.Workbooks.Open("D:/Coding/excelTest.xlsx");
            //get the first paage of the spreadsheet
            Excel._Worksheet wkSheet = wkbook.Worksheets[1];
            //find the range of used cells.
            Excel.Range usedRange = wkSheet.UsedRange;
            
            int rowCounter = usedRange.Rows.Count;
            int colCounter = usedRange.Columns.Count;

            int rowCount = 2;
            int colCount = 1;

            while (rowCount < rowCounter)
            {
               
                    items.Add(new MyItems { Thingy = usedRange.Cells[rowCount, colCount].Value2.ToString(), Thingy2 = usedRange.Cells[rowCount, colCount + 1].Value2.ToString() });
                   
                
                
                rowCount++;

            }


            wkbook.Close();
            exclApp.Quit();



        }


        public class MyItems
        {

            public string Thingy { get; set; }
            public string Thingy2 { get; set; }


        }


    }
}

