using For_Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Timetable_Solver
{
    public partial class Form1 : Form
    {
        //Create Table
        DataTable dtS7 = new DataTable();
        DataTable dtS3 = new DataTable();
        DataTable dtS2 = new DataTable();
        DataTable dtS6 = new DataTable();
        DataTable dtTime = new DataTable();
        //Define Operation time
        DateTime[] operate_start = new[] 
        { 
            DateTime.ParseExact("10:00", "HH:mm", CultureInfo.InvariantCulture),
            DateTime.ParseExact("11:00", "HH:mm", CultureInfo.InvariantCulture),
            DateTime.ParseExact("10:00", "HH:mm", CultureInfo.InvariantCulture),
            DateTime.ParseExact("10:00", "HH:mm", CultureInfo.InvariantCulture)
        };
        DateTime[] operate_end = new[]
        {
            DateTime.ParseExact("22:30", "HH:mm", CultureInfo.InvariantCulture),
            DateTime.ParseExact("23:30", "HH:mm", CultureInfo.InvariantCulture),
            DateTime.ParseExact("21:30", "HH:mm", CultureInfo.InvariantCulture),
            DateTime.ParseExact("22:30", "HH:mm", CultureInfo.InvariantCulture)
        };
        /*
        DateTime operateS7_start[0] = DateTime.ParseExact("10:00", "HH:mm", CultureInfo.InvariantCulture);
        DateTime operateS7_end = DateTime.ParseExact("22:30", "HH:mm", CultureInfo.InvariantCulture);
        DateTime operateS3_start = DateTime.ParseExact("11:00", "HH:mm", CultureInfo.InvariantCulture);
        DateTime operateS3_end = DateTime.ParseExact("23:30", "HH:mm", CultureInfo.InvariantCulture);
        DateTime operateS2_start = DateTime.ParseExact("10:00", "HH:mm", CultureInfo.InvariantCulture);
        DateTime operateS2_end = DateTime.ParseExact("21:30", "HH:mm", CultureInfo.InvariantCulture);
        DateTime operateS6_start = DateTime.ParseExact("10:00", "HH:mm", CultureInfo.InvariantCulture);
        DateTime operateS6_end = DateTime.ParseExact("22:30", "HH:mm", CultureInfo.InvariantCulture);*/
        //Define using string or numeric

        string[] week = { "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN" };
        //Labour number
        int[] picknum = { 0, 0, 0, 0, 0, 0, 0 };
        //Excel
        Excel_Generate excel_Generate = new Excel_Generate();
        public Form1()
        {
            InitializeComponent();
        }
        private void txtS7Emp1_TextChanged(object sender, EventArgs e)
        {

        }

        private void buttonSend_Click(object sender, EventArgs e)
        {
            try
            {
                #region Define Table
                dtS7.Clear(); dtS3.Clear(); dtS2.Clear(); dtTime.Clear();
                dtS7.Reset();dtS3.Reset();dtS2.Reset();dtTime.Reset();
                dtS7.Columns.Add("Name");
                dtS7.Columns.Add("MON");
                dtS7.Columns.Add("TUE");
                dtS7.Columns.Add("WED");
                dtS7.Columns.Add("THU");
                dtS7.Columns.Add("FRI");
                dtS7.Columns.Add("SAT");
                dtS7.Columns.Add("SUN");
                dtS7.Columns.Add("Picktime");
                dtS3 = dtS7.Clone();
                dtS2 = dtS7.Clone();
                dtS6 = dtS7.Clone();
                dtTime.Columns.Add("Start_Time");
                dtTime.Columns.Add("End_Time");
                dtTime.Columns.Add("MON");
                dtTime.Columns.Add("TUE");
                dtTime.Columns.Add("WED");
                dtTime.Columns.Add("THU");
                dtTime.Columns.Add("FRI");
                dtTime.Columns.Add("SAT");
                dtTime.Columns.Add("SUN");
                #endregion
                //insert table by shops
                insertTable(ref dtS7, "S7");
                insertTable(ref dtS3, "S3");
                insertTable(ref dtS2, "S2");
                insertTable(ref dtS6, "S6");
                //Arrange and write a week timetable for S7
                arrageTime(dtS7,"S7");
                //Arrange and write a week timetable for S3
                arrageTime(dtS3,"S3");
                //Arrange and write a week timetable for S2
                arrageTime(dtS2,"S2");
                //Arrange and write a week timetable for S6
                arrageTime(dtS6,"S6");
                //generate Excel
                excel_Generate.generateEXCEL(dtTime, this.txtPath.Text);

            }
            catch (Exception err)
            {
                Console.WriteLine("Sysem error!" + err);
            }
        }

        //It's a temp table for the program
        public void insertTable(ref DataTable dt, string shop)
        {
            for (int i = 1; i < 11; i++)
            {
                DataRow dtRow = dt.NewRow();
                //get the input data from textbox(Employee name)
                TextBox Name = this.Controls.Find(("txt"+shop+"Emp" + i.ToString()), true).FirstOrDefault() as TextBox;
                //While read the last emp
                if (Name is null || Name.Text == "")
                {
                    continue;
                }
                dtRow["Name"] = Convert.ToString(Name.Text);
                //get the Emp available time from richtextbox
                RichTextBox txtDetail = this.Controls.Find(("richtxt" + shop + "Emp" + i.ToString()), true).FirstOrDefault() as RichTextBox;
                if (txtDetail is null || txtDetail.Text =="")
                {
                    continue;
                }
                string Emp_detail = Convert.ToString(txtDetail.Text);
                //put the different day into array
                string[] Emp_Day = Emp_detail.Split(Environment.NewLine.ToCharArray());
                //read the array and put into the row
                for (int j = 0; j < Emp_Day.Length; j++)
                {
                    string day_detail = Emp_Day[j].Substring(4, Emp_Day[j].Length - 4);
                    //If the labour is available all day, put the operate time
                    if (day_detail == "ALL")
                    {
                        switch (shop)
                        {
                            case "S7":
                                dtRow[week[j]] = operate_start[0].ToString("HH:mm") + "-" + operate_end[0].ToString("HH:mm");
                                break;
                            case "S3":
                                dtRow[week[j]] = operate_start[1].ToString("HH:mm") + "-" + operate_end[1].ToString("HH:mm");
                                break;
                            case "S2":
                                dtRow[week[j]] = operate_start[2].ToString("HH:mm") + "-" + operate_end[2].ToString("HH:mm");
                                break;
                            default:
                                dtRow[week[j]] = operate_start[3].ToString("HH:mm") + "-" + operate_end[3].ToString("HH:mm");
                                break;
                        }
                    }
                    else
                    {
                        dtRow[week[j]] = day_detail;
                    }
                }
                //Default picktime is zero
                dtRow["Picktime"] = 0;
                //add row into table
                dt.Rows.Add(dtRow);
            }
        }

        //insert the shop title
        public void insertTitle(string shop)
        {
            DataRow dr = dtTime.NewRow();
            dr["Start_Time"] = shop;
            dtTime.Rows.Add(dr);
        }

        //***pick labour for each day based on pick number
        public void arrageTime(DataTable dtRef,string shop)
        {
            //insert the shop title
            insertTitle(shop);
            //From Mon to Sun
            for (int k =0; k<7; k++)
            {
                string Day = week[k];
                DataTable pick4today = dtRef.Clone();
                DataTable dt = new DataTable();
                var results = dtRef.AsEnumerable().Where(x => (x.Field<string>("Picktime") == "0") && (x.Field<string>(Day) != "N"));
                var t = results.FirstOrDefault();
                if (t != null && t.ToString() != "System.Data.DataRowCollection")
                {
                    pick4today = results.CopyToDataTable();
                }
                var rnd = new Random();
                if (pick4today.Rows.Count < 2)
                {
                    //find the most suitable labour based on the picktime
                    for (int i = 1; i <= 7; i++)
                    {
                        results = dtRef.AsEnumerable().Where(x => (x.Field<string>("Picktime") == i.ToString()) && (x.Field<string>(Day) != "N"));
                        var q = results.FirstOrDefault();
                        //If there is no match data, read the next pick time
                        if (q == null) { continue; }
                        //redom pick the rest people from the same pick time
                        var random12 = results.CopyToDataTable().AsEnumerable().OrderBy(r => rnd.Next()).Take(2 - pick4today.Rows.Count);
                        dt = random12.CopyToDataTable();
                        if (dt.Rows.Count > 0)
                        {
                            pick4today.ImportRow(dt.Rows[0]);
                            if (dt.Rows.Count == 2)
                            {
                                pick4today.ImportRow(dt.Rows[1]);
                            }
                            //The labour is enough
                            if (pick4today.Rows.Count == 2)
                            {
                                var rowsToUpdate =
                                dtRef.AsEnumerable().Where(r => (r.Field<string>("Name") == Convert.ToString(pick4today.Rows[0]["Name"]))
                                                          || (r.Field<string>("Name") == Convert.ToString(pick4today.Rows[1]["Name"])));
                                //Update the picktime
                                foreach (var row in rowsToUpdate)
                                {
                                    row.SetField("Picktime", Convert.ToString(Convert.ToInt16(row["Picktime"]) + 1));
                                }
                                break;
                            }
                        }
                    }
                    //if not enough labour today
                    if (pick4today.Rows.Count == 1)
                    {
                        var rowsToUpdate =
                            dtRef.AsEnumerable().Where(r => r.Field<string>("Name") == Convert.ToString(pick4today.Rows[0]["Name"]));
                        //Update the picktime
                        foreach (var row in rowsToUpdate)
                        {
                            row.SetField("Picktime", Convert.ToString(Convert.ToInt16(row["Picktime"]) + 1));
                        }
                        pick4today.Rows.Add();
                        pick4today.Rows[1]["Name"] = "";
                    }
                    else if (pick4today.Rows.Count == 0)
                    {
                        pick4today.Rows.Add();
                        pick4today.Rows[0]["Name"] = "";
                        pick4today.Rows.Add();
                        pick4today.Rows[1]["Name"] = "";
                    }
                }
                else
                {
                    var random12 = pick4today.AsEnumerable().OrderBy(r => rnd.Next()).Take(2);
                    dt = random12.CopyToDataTable();
                    pick4today.Clear();
                    pick4today = dt;
                    var rowsToUpdate =
                            dtRef.AsEnumerable().Where(r => (r.Field<string>("Name") == Convert.ToString(pick4today.Rows[0]["Name"]))
                                                          || (r.Field<string>("Name") == Convert.ToString(pick4today.Rows[1]["Name"])));
                    //Update the picktime
                    foreach (var row in rowsToUpdate)
                    {
                        row.SetField("Picktime", Convert.ToString(Convert.ToInt16(row["Picktime"]) + 1));
                    }
                }

                //Write to the temp table (For excel)
                insertTime(pick4today, week[k],shop);
            }
        }

        public void insertTime(DataTable Pickfortoday, string day,string shop)
        {
            //Transfer shop to num
            int num = 0;
            switch (shop)
            {
                case "S7": 
                    num = 0;
                    break;
                case "S3":
                    num = 1;
                    break;
                case "S2":
                    num = 2;
                    break;
                case "S6":
                    num = 3;
                    break;

            }
            //When write the shift for Monday, the time will also insert
            if (day == "MON")
            {
                for (int i = 0; i < 2; i++)
                {
                    DataRow dr = dtTime.NewRow();
                    dr["Start_Time"] = operate_start[num].ToString("HH:mm");
                    dr["End_Time"] = operate_end[num].ToString("HH:mm");
                    dr[day] = Convert.ToString(Pickfortoday.Rows[i]["Name"]);
                    dtTime.Rows.Add(dr);
                }
            }
            else
            {
                dtTime.Rows[dtTime.Rows.Count-1][day] = Pickfortoday.Rows[0]["Name"];
                dtTime.Rows[dtTime.Rows.Count-2][day] = Pickfortoday.Rows[1]["Name"];
            }

        }

        //Select File Path
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            this.txtPath.Text = path.SelectedPath;
        }
    }
}
