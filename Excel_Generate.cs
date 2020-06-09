using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace For_Excel
{
    public class Excel_Generate
    {

        public void generateEXCEL(DataTable dt,string Path)
        {
            //set the location
            string FileStr = Path+"\\Schedule";
            //Application
            Excel.Application Excel_App1 = new Excel.Application();
            //File
            Excel.Workbook Excel_WB1 = Excel_App1.Workbooks.Add();
            //Define Sheet
            Excel.Worksheet Excel_WS1 = new Excel.Worksheet();
            Excel_WS1 = Excel_WB1.Worksheets[1];
            writeExcel(Excel_WS1, dt);

            //save
            Excel_WB1.SaveAs(FileStr);

            //close
            Excel_WS1 = null;
            Excel_WB1.Close();
            Excel_WB1 = null;
            Excel_App1.Quit();
            Excel_App1 = null;

        }

        public void writeExcel(Excel.Worksheet wb, DataTable dt)
        {
            if (dt.Rows.Count > 0)
            {
                wb.Name = "Shift_Table";
                writeTitle(wb, dt);
                int Row = dt.Rows.Count;
                int Col = dt.Columns.Count;
                for (int i = 0; i < Row; i++)
                {
                    for (int j = 0; j < Col; j++)
                    {
                        string str = dt.Rows[i][j].ToString();
                        wb.Cells[i + 2, j+1] = str;
                    }
                }
            }
        }

        private void writeTitle(Excel.Worksheet wb,DataTable dt)
        {
            int col = dt.Columns.Count;
            for(int i =3; i <= col; i++)
            {
                wb.Cells[1, i] = dt.Columns[i-1].ColumnName;
            }
        }


    }
    
}
