using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Sql;
using System.Data.Common;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.IO;
using System.Threading;

namespace ReadCashFlow
{
    public partial class Form1 : Form
    {
        public Form1(string[] param)
        {
            InitializeComponent();

            label1.Visible = false;
            if (!string.IsNullOrEmpty(param[2]) && !string.IsNullOrEmpty(param[0]))
            {
                label1.Visible = true;
                //this.Show();
                //MessageBox.Show("before thread");
                Thread newThread = new Thread(DoWork);

                object args = new object[5] { param[0], param[1], param[2], param[3], param[4] };
                //MessageBox.Show("before start thread");
                newThread.Start(args);
            }
        }
        public void DoWork(object args)
        {
            Array argArray = new object[5];
            argArray = (Array)args;
            string[] param = new string[5];
            param[0] = (string)argArray.GetValue(0);
            param[1] = (string)argArray.GetValue(1);
            param[2] = (string)argArray.GetValue(2);
            param[3] = (string)argArray.GetValue(3);
            param[4] = (string)argArray.GetValue(4);

            SqlTransaction trans = null;
            string _PYMDBconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PCMSConnectionString"].ConnectionString;

            //DirectoryInfo di = new DirectoryInfo(Environment.CurrentDirectory);
            DirectoryInfo di = new DirectoryInfo(param[2]);
            FileInfo[] rgFiles = di.GetFiles("*.xls");

            if (!string.IsNullOrEmpty(param[2]))
                if (File.Exists(param[2] + "\\Result.html"))
                    File.Delete(param[2] + "\\Result.html");

            Dictionary<string, string> error_sheet = new Dictionary<string, string>();
            FileStream file2 = new FileStream(param[2] + "\\Result.html", FileMode.Append);

            FileInfo[] error_Files = di.GetFiles("*.txt");
            foreach (FileInfo error_f in error_Files)
            {
                if (error_f.Name.StartsWith("error"))
                    if (File.Exists(error_f.FullName))
                        File.Delete(error_f.FullName);
            }
            //FileStream file = new FileStream(Environment.CurrentDirectory + "\\error.txt", FileMode.Append);
           
            Dictionary<String, String> log_full_name = new Dictionary<String, String>();
            Dictionary<String, String> log_sheet_name = new Dictionary<String, String>();
            Dictionary<String, String> validation_status = new Dictionary<String, String>();
            Dictionary<String, String> upload_status = new Dictionary<String, String>();
            Dictionary<String, String> error_status = new Dictionary<String, String>();
            using (StreamWriter sw2 = new StreamWriter(file2))
            {
                sw2.WriteLine("<html>");
                sw2.WriteLine("<table border=1>");
                sw2.WriteLine(" <tr>");
                sw2.WriteLine("     <td width=50>");
                sw2.WriteLine("         File Name");
                sw2.WriteLine("     </td>");
                sw2.WriteLine("     <td>");
                sw2.WriteLine("         Worksheet");
                sw2.WriteLine("     </td>");
                sw2.WriteLine("     <td>");
                sw2.WriteLine("         Validation Status");
                sw2.WriteLine("     </td>");
                sw2.WriteLine("     <td>");
                sw2.WriteLine("         Upload Status");
                sw2.WriteLine("     </td>");
                sw2.WriteLine("     <td>");
                sw2.WriteLine("         Error");
                sw2.WriteLine("     </td>");
                sw2.WriteLine(" </tr>");
                // Display the ProgressBar control.
                progressBar1.Visible = true;
                // Set Minimum to 1 to represent the first file being copied.
                progressBar1.Minimum = 1;
                // Set Maximum to the total number of files to copy.
                progressBar1.Maximum = rgFiles.Count() + 10;
                // Set the initial value of the ProgressBar.
                progressBar1.Value = 1;
                // Set the Step property to a value of 1 to represent each file being copied.
                progressBar1.Step = 1;

                foreach (FileInfo fi in rgFiles)
                {
                    //Console.WriteLine(param[2] + param[3]);
                    //Console.WriteLine("wait");
                    //Console.ReadKey();

                    label1.Text = fi.Name;
                    progressBar1.PerformStep();
                    //if ((param[2] + param[3]) == fi.FullName && param[0].ToUpper() == "ONE")
                    //if ((param[2] + "\\" + param[3]) == fi.FullName && param[0].ToUpper() == "ONE") //for Console
                    if ((param[2] + param[3]) == fi.FullName && param[0].ToUpper() == "ONE") // for Winform
                    {
                        #region one
                        object Missing = System.Type.Missing;
                        Excel._Workbook oWB = null;
                        Excel._Worksheet oSheet = null;
                        Excel.Range oRng = null;
                        Excel.Application oXL = new Excel.Application();
                        try
                        {
                            oXL.Visible = false;
                            oXL.DisplayAlerts = false;

                            //Get a new workbook.
                            //oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing));

                            oWB = oXL.Workbooks.Open(fi.FullName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing);

                            string CashFlow_Header_ID = "";
                            bool once_header = false;
                            string Budget_Project_No = "";
                            bool once_worksheet = false;

                            // validation

                            string error_filename = "error_" + param[3] + "_" + DateTime.Now.ToString().Replace(":", "-").Replace("/", "-");
                            string error_path = param[2] + "\\" + error_filename + ".txt";
                            FileStream file = new FileStream(error_path, FileMode.Append);
                            //FileStream file = new FileStream(Environment.CurrentDirectory + "\\error.txt", FileMode.Append);
                            bool isWrittenEndHtml = false;
                            using (StreamWriter sw = new StreamWriter(file))
                            {

                                for (int i = 1; i <= oWB.Sheets.Count; i++)
                                {
                                    //oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                                    oSheet = (Excel._Worksheet)oWB.Sheets[i];


                                    Boolean valid = true;
                                    // Validate Header
                                    string[] types = new string[2];
                                    types[0] = "BUDGET";
                                    types[1] = "CASHFLOW";
                                    once_worksheet = false;
                                    string Worksheet_ID = "";
                                    int number_of_error = 0;
                                    bool user_custom_sheet = false;

                                    int counter2 = 0;
                                    foreach (string type in types)
                                    {
                                        counter2 = counter2 + 1;
                                        string latestperiod = LatestPeriodInReport();
                                        //string startrow = StartRow(type, latestperiod);

                                        //select * from dbo.Budget_CashFlow_Excel_ReportItem_Detail where Heading_Type = 'D'
                                        Dictionary<string, string[]> BudgetHeaderDict = ReadHeaderFromDB(type);
                                        foreach (string k in BudgetHeaderDict.Keys)
                                        {
                                            string[] location = BudgetHeaderDict[k];
                                            oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);

                                        }

                                        oRng = oSheet.get_Range("A1", "A1");

                                        if (oRng.Value2 == null)
                                        {
                                            user_custom_sheet = true;
                                            number_of_error = number_of_error + 1;
                                        }
                                        else
                                            if (string.IsNullOrEmpty(oRng.Value2.ToString()))
                                                number_of_error = number_of_error + 1;

                                        if (oRng.Value2 != null)
                                        {
                                            if (!string.IsNullOrEmpty(oRng.Value2.ToString()))
                                            {
                                                if (counter2 == 1)
                                                {
                                                    log_full_name.Add(oSheet.Name, fi.FullName);
                                                    log_sheet_name.Add(oSheet.Name, oSheet.Name);
                                                    /*
                                                    sw2.WriteLine(" <tr>");
                                                    sw2.WriteLine("     <td width=50>");
                                                    sw2.WriteLine("         " + fi.FullName);
                                                    sw2.WriteLine("     </td>");
                                                    sw2.WriteLine("     <td>");
                                                    sw2.WriteLine("         " + oSheet.Name);
                                                    sw2.WriteLine("     </td>");
                                                    */
                                                }
                                                int year_end_index = oRng.Value2.ToString().IndexOf("The Year Ended");
                                                //string year_end_date_string = oRng.Value2.ToString().Substring(year_end_index, oRng.Value2.ToString().Length - year_end_index).Replace("The Year Ended", "").Trim();
                                                //string year_end_date_string = ;// buster date
                                                //DateTime year_end_date = Convert.ToDateTime(year_end_date_string);
                                                DateTime year_end_date = Convert.ToDateTime("1-" + (Convert.ToInt32(param[4].Substring(4, 2)) + 1).ToString() + "-" + param[4].Substring(0, 4)).AddDays(-1);

                                                string BudgetWorkSheetSQL = "Insert into ";
                                                string BudgetWorkSheetSQL_FieldName = "";
                                                string BudgetWorkSheetSQL_Value = "";

                                                string[] Budget_Project_On_Hand_Status_location = ReadBudget_Project_On_Hand_StatusFromDB(latestperiod);
                                                oRng = oSheet.get_Range(Budget_Project_On_Hand_Status_location[1] + Budget_Project_On_Hand_Status_location[0], Budget_Project_On_Hand_Status_location[1] + Budget_Project_On_Hand_Status_location[0]);
                                                string Budget_Project_On_Hand_Status = oRng.Value2.ToString();


                                                Dictionary<string, string[]> BudgetHeaderItemRowColDict = ReadHeaderItemRowColFromDB(Budget_Project_On_Hand_Status, latestperiod);
                                                Dictionary<string[], string[]> BudgetDetailDict = ReadDetailFromDB(type, Budget_Project_On_Hand_Status);
                                                Dictionary<string, string[]> BudgetDetailColDict = ReadDetailColFromDB(type);
                                                Dictionary<string, string[]> Budget_CashFlow_ReportItem_CodeDict = ReadBudget_CashFlow_ReportItem_CodeFromDB(type);
                                                Dictionary<string, string[]> Budget_CashFlow_AllReportItem_CodeDict = ReadBudget_CashFlow_AllReportItem_CodeFromDB(type);


                                                string _connectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;

                                                if (once_worksheet == false)
                                                {
                                                    using (SqlConnection connection = new SqlConnection(_connectionString))
                                                    {
                                                        try
                                                        {
                                                            int id = 0;

                                                            foreach (string k in BudgetHeaderItemRowColDict.Keys)
                                                            {
                                                                string[] location = new string[2];
                                                                BudgetHeaderItemRowColDict.TryGetValue(k, out location);
                                                                oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);


                                                                if ((k.Contains("YrMth") || k.Contains("Date") || k.Contains("date")))
                                                                {
                                                                    /*
                                                                    DateTime datetime_parse_result;
                                                                    
                                                                    if (DateTime.TryParse(oRng.get_Value(Type.Missing).ToString().Replace("After","").Replace(" ","").Trim(), out datetime_parse_result) == true)
                                                                    {
                                                                        //command.Parameters.Add(new SqlParameter("@" + k, Convert.ToDateTime(oRng.get_Value(Type.Missing)).ToString("yyyy/MM/dd HH:mm:ss")));
                                                                        Console.Write("");
                                                                    }
                                                                    else
                                                                    {
                                                                        number_of_error = number_of_error + 1;
                                                                        sw.WriteLine(oSheet.Name + ": " + k + " is not a date");
                                                                    }*/
                                                                }
                                                                else
                                                                    if (oRng.Value2 == null)
                                                                    {
                                                                        Console.Write("");
                                                                    }
                                                                    else
                                                                    {
                                                                        decimal decimal_out = 0;
                                                                        if (decimal.TryParse(oRng.Value2.ToString().Replace("'", "''"), out decimal_out) == true)
                                                                        {
                                                                            if (k.Contains("Budget_Project_No"))
                                                                                Budget_Project_No = oRng.Value2.ToString().Replace("'", "''");
                                                                        }
                                                                        else
                                                                            if (k.Contains("Hit_Rate"))
                                                                            {
                                                                                Console.Write("");
                                                                                //command.Parameters.Add(new SqlParameter("@" + k, decimal_out));
                                                                            }
                                                                            else
                                                                            {
                                                                                if (k.Contains("Budget_Project_No"))
                                                                                    Budget_Project_No = oRng.Value2.ToString().Replace("'", "''");
                                                                            }
                                                                    }
                                                            }
                                                            //BudgetWorkSheetSQL = "insert into Budget_CashFlow_Worksheet (" + BudgetWorkSheetSQL_FieldName + ") values (" + BudgetWorkSheetSQL_Value + ")";
                                                            //ExecuteDatabase(BudgetWorkSheetSQL);

                                                            once_worksheet = true;
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            connection.Close();
                                                        }
                                                    }
                                                    string number_of_found = ProjectCodeExistInPCMS(Budget_Project_No);
                                                    if (string.IsNullOrEmpty(number_of_found) && Budget_Project_On_Hand_Status != "Unsecured")
                                                    {
                                                        number_of_error = number_of_error + 1;
                                                        sw.WriteLine(oSheet.Name + ": Project Code/No not exist in PCMS");
                                                    }
                                                    else
                                                    {
                                                        int num = 0;
                                                        if (Int32.TryParse(number_of_found, out num) == false && Budget_Project_On_Hand_Status != "Unsecured")
                                                        {
                                                            number_of_error = number_of_error + 1;
                                                            sw.WriteLine(oSheet.Name + ": Project Code/No not exist in PCMS");
                                                        }
                                                        if (num == 0 && Budget_Project_On_Hand_Status != "Unsecured")
                                                        {
                                                            number_of_error = number_of_error + 1;
                                                            sw.WriteLine(oSheet.Name + ": Project Code/No not exist in PCMS");
                                                        }
                                                    }
                                                    //if (oSheet.Name != Budget_Project_No)
                                                    //sw.WriteLine(k[1] + " Missing Report Code");
                                                }
                                                string Budget_CashFlow_Project_ID = "";
                                                string BudgetDetailSQL = "Insert into ";
                                                string BudgetDetailSQL_FieldName = "";
                                                string BudgetDetailSQL_Value = "";
                                                // Insert into
                                                string ErrorReport = "";
                                                try
                                                {
                                                    if (!string.IsNullOrEmpty(Budget_Project_No))
                                                    {
                                                        int counter = 0;
                                                        foreach (string[] k in BudgetDetailDict.Keys)
                                                        {
                                                            if (counter == 8)
                                                                Console.WriteLine("");
                                                            counter = counter + 1;
                                                            //id[0] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                                            //id[1] = dr["Budget_CashFlow_ReportItem_Name"].ToString();
                                                            //value[0] = dr["Row_No"].ToString();
                                                            //value[1] = dr["Budget_CashFlow_ReportItem_ID"].ToString();
                                                            //value[2] = dr["Budget_CashFlow_ReportItem_Code"].ToString();

                                                            string[] row = BudgetDetailDict[k];

                                                            if (Worksheet_ID == "1")
                                                                Console.Write("");
                                                            if (row[2].Replace("'", "''") == "COMPULSORY" || row[2].Replace("'", "''") == "WARNING")
                                                            {
                                                                foreach (string k2 in Budget_CashFlow_ReportItem_CodeDict.Keys)
                                                                {
                                                                    if (k2 == "Budget_CashFlow_ReportItem_Code") // i.e. get column C
                                                                    {
                                                                        string[] data = new string[2];
                                                                        Budget_CashFlow_ReportItem_CodeDict.TryGetValue(k2, out data);
                                                                        string col = data[0];

                                                                        if (!string.IsNullOrEmpty(col))
                                                                        {
                                                                            oRng = oSheet.get_Range(col + row[0], col + row[0]);
                                                                            if (Worksheet_ID == "1")
                                                                                Console.Write("");
                                                                            if (oRng.Value2 != null)
                                                                            {
                                                                                if (!string.IsNullOrEmpty(data[1]))
                                                                                {
                                                                                    /*
                                                                                    switch (data[1])
                                                                                    {
                                                                                        case "C":
                                                                                            DateTime datetime_result2;
                                                                                            if (DateTime.TryParse(oRng.Value2.ToString(), out datetime_result2) == true)
                                                                                            {
                                                                                                number_of_error = number_of_error + 1;
                                                                                                sw.WriteLine(oSheet.Name + ":It is not a String - " + k[1]);
                                                                                            }
                                                                                            decimal numeric_result2 = 0;
                                                                                            if (Decimal.TryParse(oRng.Value2.ToString(), out numeric_result2) == true)
                                                                                            {
                                                                                                number_of_error = number_of_error + 1;
                                                                                                sw.WriteLine(oSheet.Name + ":It is not a String - " + k[1]);
                                                                                            }
                                                                                            break;
                                                                                        case "D":
                                                                                            DateTime datetime_result;
                                                                                            if (DateTime.TryParse(oRng.Value2.ToString(), out datetime_result) == false)
                                                                                            {
                                                                                                number_of_error = number_of_error + 1;
                                                                                                sw.WriteLine(oSheet.Name + ":It is not a datetime format - " + k[1]);
                                                                                            }
                                                                                            break;
                                                                                        case "N":
                                                                                            decimal numeric_result = 0;
                                                                                            if (Decimal.TryParse(oRng.Value2.ToString(), out numeric_result) == false)
                                                                                            {
                                                                                                number_of_error = number_of_error + 1;
                                                                                                sw.WriteLine(oSheet.Name + ":It is not numeric - " + k[1]);
                                                                                            }
                                                                                            break;
                                                                                        default:
                                                                                            break;
                                                                                    }*/
                                                                                }
                                                                                if (type == "BUDGET")
                                                                                {
                                                                                    string number_of_found = CostCodeExistInPCMS(oRng.Value2.ToString().Replace("'", "''"));
                                                                                    if (string.IsNullOrEmpty(number_of_found))
                                                                                    {
                                                                                        number_of_error = number_of_error + 1;
                                                                                        sw.WriteLine(oSheet.Name + ":Missing Report Code at " + col + row[0] + " - " + k[1]);
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        int num = 0;
                                                                                        if (Int32.TryParse(number_of_found, out num) == false)
                                                                                        {
                                                                                            number_of_error = number_of_error + 1;
                                                                                            sw.WriteLine(oSheet.Name + ":Missing CashFlow Code and number of found is not an integer " + col + row[0] + " - " + k[1]);
                                                                                        }
                                                                                        if (num == 0)
                                                                                        {
                                                                                            number_of_error = number_of_error + 1;
                                                                                            sw.WriteLine(oSheet.Name + ":Missing Report Code " + col + row[0] + " - " + k[1]);
                                                                                        }
                                                                                    }
                                                                                }
                                                                                else if (type == "CASHFLOW")
                                                                                {
                                                                                    string number_of_found = CashFlowCodeExistInPCMS(oRng.Value2.ToString().Replace("'", "''"));
                                                                                    if (string.IsNullOrEmpty(number_of_found))
                                                                                    {
                                                                                        number_of_error = number_of_error + 1;
                                                                                        sw.WriteLine(oSheet.Name + ":Missing CashFlow Code " + col + row[0] + " - " + k[1]);
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        int num = 0;
                                                                                        if (Int32.TryParse(number_of_found, out num) == false)
                                                                                        {
                                                                                            number_of_error = number_of_error + 1;
                                                                                            sw.WriteLine(oSheet.Name + ":Missing CashFlow Code and number of found is not an integer " + col + row[0] + " - " + k[1]);
                                                                                        }
                                                                                        if (num == 0)
                                                                                        {
                                                                                            number_of_error = number_of_error + 1;
                                                                                            sw.WriteLine(oSheet.Name + ":Missing CashFlow Code " + col + row[0] + " - " + k[1]);
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                if ((row[2].Replace("'", "''") == "COMPULSORY" || row[2].Replace("'", "''") == "WARNING"))
                                                                                {
                                                                                    if (type == "BUDGET")
                                                                                    {
                                                                                        number_of_error = number_of_error + 1;
                                                                                        sw.WriteLine(oSheet.Name + ":Missing Report Code " + col + row[0] + " - " + k[1]);
                                                                                    }
                                                                                    else if (type == "CASHFLOW")
                                                                                    {
                                                                                        number_of_error = number_of_error + 1;
                                                                                        sw.WriteLine(oSheet.Name + ":Missing CashFlow Code " + col + row[0] + " - " + k[1]);
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            else
                                                                BudgetDetailSQL_Value = "'" + row[2].Replace("'", "''") + "'";

                                                        }
                                                        //validate header
                                                        foreach (string k3 in BudgetHeaderItemRowColDict.Keys)
                                                        {
                                                            string[] location = new string[3];
                                                            BudgetHeaderItemRowColDict.TryGetValue(k3, out location);
                                                            oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);

                                                            switch (location[2])
                                                            {
                                                                case "D":
                                                                    DateTime datetime_result;
                                                                    if (oRng.get_Value(Missing) != null)
                                                                    {
                                                                        if (!oRng.get_Value(Missing).ToString().Replace("After", "").Contains("N/A"))
                                                                        {
                                                                            if (DateTime.TryParse(oRng.get_Value(Missing).ToString().Replace("After", ""), out datetime_result) == false)
                                                                            {
                                                                                number_of_error = number_of_error + 1;
                                                                                sw.WriteLine(oSheet.Name + ":It is not a datetime format at " + location[1] + location[0] + " - " + k3);
                                                                            }
                                                                        }
                                                                    }
                                                                    break;
                                                                case "N":
                                                                    decimal numeric_result = 0;
                                                                    if (oRng.get_Value(Missing) != null)
                                                                    {
                                                                        // buster 30 Jan
                                                                        if (!oRng.get_Value(Missing).ToString().Replace("After", "").Contains("N/A"))
                                                                        {
                                                                            if (Decimal.TryParse(oRng.get_Value(Missing).ToString(), out numeric_result) == false)
                                                                            {
                                                                                number_of_error = number_of_error + 1;
                                                                                sw.WriteLine(oSheet.Name + ":It is not numeric at " + location[1] + location[0] + " - " + k3);
                                                                                //if ("H5" == location[1] + location[0])
                                                                                //Console.WriteLine("");
                                                                            }
                                                                        }
                                                                    }
                                                                    break;
                                                                default:
                                                                    break;
                                                            }
                                                        }
                                                        //validate data
                                                        foreach (string[] k in BudgetDetailDict.Keys)
                                                        {
                                                            //id[0] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                                            //id[1] = dr["Budget_CashFlow_ReportItem_Name"].ToString();
                                                            //value[0] = dr["Row_No"].ToString();
                                                            //value[1] = dr["Budget_CashFlow_ReportItem_ID"].ToString();
                                                            //value[2] = dr["Budget_CashFlow_ReportItem_Code"].ToString();

                                                            string[] row = BudgetDetailDict[k];

                                                            foreach (string k2 in Budget_CashFlow_AllReportItem_CodeDict.Keys)
                                                            {
                                                                string[] data = new string[2];
                                                                Budget_CashFlow_AllReportItem_CodeDict.TryGetValue(k2, out data);
                                                                string col = data[0];

                                                                if (!string.IsNullOrEmpty(col))
                                                                {
                                                                    oRng = oSheet.get_Range(col + row[0], col + row[0]);

                                                                    switch (data[1])
                                                                    {
                                                                        case "D":
                                                                            DateTime datetime_result;
                                                                            if (oRng.get_Value(Missing) != null)
                                                                            {
                                                                                if (DateTime.TryParse(oRng.get_Value(Missing).ToString(), out datetime_result) == false)
                                                                                {
                                                                                    number_of_error = number_of_error + 1;
                                                                                    sw.WriteLine(oSheet.Name + ":It is not a datetime format at " + col + row[0] + " - " + k[1]);
                                                                                }
                                                                            }
                                                                            break;
                                                                        case "N":
                                                                            decimal numeric_result = 0;
                                                                            if (oRng.get_Value(Missing) != null)
                                                                            {
                                                                                if (Decimal.TryParse(oRng.get_Value(Missing).ToString(), out numeric_result) == false)
                                                                                {
                                                                                    number_of_error = number_of_error + 1;
                                                                                    sw.WriteLine(oSheet.Name + ":It is not numeric at " + col + row[0] + " - " + k[1]);
                                                                                }
                                                                            }
                                                                            break;
                                                                        default:
                                                                            break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine(ex.Message.ToString());
                                                }
                                            }
                                        }
                                    }
                                    if (user_custom_sheet == false)
                                    {
                                        if (number_of_error > 0)
                                        {
                                            isWrittenEndHtml = true;

                                            validation_status.Add(oSheet.Name, "FAILED");
                                            upload_status.Add(oSheet.Name, "FAILED");
                                            error_status.Add(oSheet.Name, "         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            /*
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         FAILED");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         FAILED");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine(" </tr>");
                                            */
                                            error_sheet.Add(error_sheet.Count.ToString(), oSheet.Name);
                                        }
                                        else
                                        {
                                            //isWrittenEndHtml = true;
                                            validation_status.Add(oSheet.Name, "SUCCESS");
                                            /*
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         SUCCESS");
                                            sw2.WriteLine("     </td>");
                                            */
                                            //sw2.WriteLine("     <td>");
                                            //sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            //sw2.WriteLine("     </td>");
                                            //sw2.WriteLine(" </tr>");
                                        }
                                    }
                                }

                                if (error_sheet.Count == 0 && !string.IsNullOrEmpty(Budget_Project_No))
                                {
                                    // process
                                    for (int i = 1; i <= oWB.Sheets.Count; i++)
                                    {
                                        //oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                                        oSheet = (Excel._Worksheet)oWB.Sheets[i];

                                        Boolean valid = true;
                                        // Validate Header
                                        string[] types = new string[2];
                                        types[0] = "BUDGET";
                                        types[1] = "CASHFLOW";
                                        once_worksheet = false;
                                        string Worksheet_ID = "";



                                        foreach (string type in types)
                                        {
                                            string latestperiod = LatestPeriodInReport();
                                            //string startrow = StartRow(type, latestperiod);

                                            //select * from dbo.Budget_CashFlow_Excel_ReportItem_Detail where Heading_Type = 'D'
                                            Dictionary<string, string[]> BudgetHeaderDict = ReadHeaderFromDB(type);
                                            foreach (string k in BudgetHeaderDict.Keys)
                                            {
                                                string[] location = BudgetHeaderDict[k];
                                                oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);

                                            }
                                            /*
                                            startrow = StartRow("CASHFLOW", LatestPeriodInReport());
                    
                                            Dictionary<string, string[]> CashFlowHeaderDict = ReadHeaderFromDB("CASHFLOW");
                                            foreach (string k in CashFlowHeaderDict.Keys)
                                            {
                                                string[] location = BudgetHeaderDict[k];
                                                oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);
                                            }*/
                                            int number_of_error = 0;
                                            oRng = oSheet.get_Range("A1", "A1");
                                            if (oRng.Value2 == null)
                                                number_of_error = number_of_error + 1;
                                            else
                                                if (string.IsNullOrEmpty(oRng.Value2.ToString()))
                                                    number_of_error = number_of_error + 1;

                                            if (oRng.Value2 != null)
                                            {
                                                if (!string.IsNullOrEmpty(oRng.Value2.ToString()))
                                                {

                                                    int year_end_index = oRng.Value2.ToString().IndexOf("The Year Ended");
                                                    //string year_end_date_string = oRng.Value2.ToString().Substring(year_end_index, oRng.Value2.ToString().Length - year_end_index).Replace("The Year Ended", "").Trim();
                                                    //DateTime year_end_date = Convert.ToDateTime(year_end_date_string);
                                                    DateTime year_end_date = Convert.ToDateTime("1-" + (Convert.ToInt32(param[4].Substring(4, 2)) + 1).ToString() + "-" + param[4].Substring(0, 4)).AddDays(-1);

                                                    string BudgetWorkSheetSQL = "Insert into ";
                                                    string BudgetWorkSheetSQL_FieldName = "";
                                                    string BudgetWorkSheetSQL_Value = "";

                                                    string[] Budget_Project_On_Hand_Status_location = ReadBudget_Project_On_Hand_StatusFromDB(latestperiod);
                                                    oRng = oSheet.get_Range(Budget_Project_On_Hand_Status_location[1] + Budget_Project_On_Hand_Status_location[0], Budget_Project_On_Hand_Status_location[1] + Budget_Project_On_Hand_Status_location[0]);
                                                    string Budget_Project_On_Hand_Status = oRng.Value2.ToString();


                                                    Dictionary<string, string[]> BudgetHeaderItemRowColDict = ReadHeaderItemRowColFromDB(Budget_Project_On_Hand_Status, latestperiod);
                                                    Dictionary<string[], string[]> BudgetDetailDict = ReadDetailFromDB(type, Budget_Project_On_Hand_Status);
                                                    Dictionary<string, string[]> BudgetDetailColDict = ReadDetailColFromDB(type);
                                                    Dictionary<string, string[]> Budget_CashFlow_ReportItem_CodeDict = ReadBudget_CashFlow_ReportItem_CodeFromDB(type);

                                                    string _connectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
                                                    if (once_header == false)
                                                    {
                                                        using (SqlConnection connection = new SqlConnection(_connectionString))
                                                        {
                                                            try
                                                            {
                                                                SqlTransaction transa = null;
                                                                SqlCommand command = new SqlCommand("sp_Budget_CashFlow_Header_Update", connection);
                                                                command.CommandType = CommandType.StoredProcedure;

                                                                int id = 0;
                                                                SqlParameter returnValue2 = new SqlParameter("@Budget_CashFlow_Header_ID", id); // ok
                                                                returnValue2.Direction = ParameterDirection.InputOutput;

                                                                command.Parameters.Add(returnValue2);

                                                                command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Upload_By", param[1]));

                                                                SqlParameter returnValue = new SqlParameter("@Result", ""); // ok
                                                                returnValue.Direction = ParameterDirection.Output;

                                                                command.Parameters.Add(returnValue);

                                                                connection.Open();
                                                                transa = connection.BeginTransaction();
                                                                command.Transaction = transa;
                                                                command.Connection = connection;
                                                                command.ExecuteNonQuery();
                                                                transa.Commit();

                                                                CashFlow_Header_ID = command.Parameters["@Budget_CashFlow_Header_ID"].Value.ToString();
                                                                string result = command.Parameters["@Result"].Value.ToString();

                                                                command.Dispose();
                                                                transa.Dispose();
                                                                connection.Close();
                                                                once_header = true;
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                sw.WriteLine(oSheet.Name + ": Stored Proc sp_Budget_CashFlow_Header_Update error " + ex.Message.ToString());
                                                                number_of_error = number_of_error + 1;
                                                                connection.Close();
                                                            }
                                                        }
                                                    }

                                                    if (once_worksheet == false && number_of_error == 0)
                                                    {
                                                        using (SqlConnection connection = new SqlConnection(_connectionString))
                                                        {
                                                            try
                                                            {
                                                                SqlTransaction transa = null;
                                                                SqlCommand command = new SqlCommand("sp_Budget_CashFlow_Worksheet_Update", connection);
                                                                command.CommandType = CommandType.StoredProcedure;

                                                                int id = 0;
                                                                SqlParameter returnValue2 = new SqlParameter("@Budget_CashFlow_Worksheet_ID", id); // ok
                                                                returnValue2.Direction = ParameterDirection.InputOutput;

                                                                command.Parameters.Add(returnValue2);

                                                                //command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Worksheet_ID", "0"));
                                                                command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Header_ID", CashFlow_Header_ID));
                                                                command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Worksheet_Name", oSheet.Name));
                                                                command.Parameters.Add(new SqlParameter("@YrEnd_Date", year_end_date.ToString("yyyy/MM/dd HH:mm:ss")));
                                                                command.Parameters.Add(new SqlParameter("@Budget_Project_On_Hand_Status", Budget_Project_On_Hand_Status));
                                                                command.Parameters.Add(new SqlParameter("@Upload_FileName", fi.FullName));


                                                                SqlParameter returnValue = new SqlParameter("@Result", ""); // ok
                                                                returnValue.Direction = ParameterDirection.Output;

                                                                command.Parameters.Add(returnValue);

                                                                //command.Parameters.Add(new SqlParameter("@Result", ""));


                                                                foreach (string k in BudgetHeaderItemRowColDict.Keys)
                                                                {
                                                                    /*
                                                                    if (string.IsNullOrEmpty(BudgetWorkSheetSQL_FieldName))
                                                                        BudgetWorkSheetSQL_FieldName = k;
                                                                    else
                                                                        BudgetWorkSheetSQL_FieldName = BudgetWorkSheetSQL_FieldName + "," + k;
                                                                    */
                                                                    string[] location = new string[2];
                                                                    BudgetHeaderItemRowColDict.TryGetValue(k, out location);
                                                                    oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);

                                                                    if (location[1] + location[0] == "AM10")
                                                                        Console.WriteLine("");

                                                                    if ((k.Contains("YrMth") || k.Contains("Date") || k.Contains("date")))
                                                                        if (oRng.get_Value(Type.Missing).GetType() == System.Type.GetType("System.DateTime"))
                                                                            command.Parameters.Add(new SqlParameter("@" + k, Convert.ToDateTime(oRng.get_Value(Type.Missing)).ToString("yyyy/MM/dd HH:mm:ss")));
                                                                        else
                                                                        {
                                                                            if (string.IsNullOrEmpty(oRng.get_Value(Type.Missing).ToString()))
                                                                                command.Parameters.Add(new SqlParameter("@" + k, System.DBNull.Value));
                                                                            else
                                                                                if (oRng.get_Value(Type.Missing).ToString() == "N/A")
                                                                                    command.Parameters.Add(new SqlParameter("@" + k, System.DBNull.Value));
                                                                                else
                                                                                    command.Parameters.Add(new SqlParameter("@" + k, oRng.get_Value(Type.Missing).ToString()));
                                                                        }
                                                                    else
                                                                        if (oRng.Value2 == null)
                                                                            command.Parameters.Add(new SqlParameter("@" + k, System.DBNull.Value));
                                                                        else
                                                                        {
                                                                            decimal decimal_out = 0;
                                                                            if (decimal.TryParse(oRng.Value2.ToString().Replace("'", "''"), out decimal_out) == true)
                                                                            {
                                                                                if (k.Contains("Budget_Project_No"))
                                                                                    Budget_Project_No = oRng.Value2.ToString().Replace("'", "''");
                                                                                command.Parameters.Add(new SqlParameter("@" + k, decimal_out));
                                                                            }
                                                                            else
                                                                                if (k.Contains("Hit_Rate"))
                                                                                    command.Parameters.Add(new SqlParameter("@" + k, decimal_out));
                                                                                else
                                                                                {
                                                                                    if (k.Contains("Budget_Project_No"))
                                                                                        Budget_Project_No = oRng.Value2.ToString().Replace("'", "''");
                                                                                    command.Parameters.Add(new SqlParameter("@" + k, oRng.Value2.ToString().Replace("'", "''")));
                                                                                }
                                                                        }

                                                                    //if (k == "CumTo_ForPeriod_YrMth") // debug
                                                                    //Console.WriteLine("");
                                                                    /*

                                                                    if (string.IsNullOrEmpty(BudgetWorkSheetSQL_Value))
                                                                        if ((k.Contains("YrMth") || k.Contains("Date") || k.Contains("date")))
                                                                            if (oRng.get_Value(Type.Missing).GetType() == System.Type.GetType("System.DateTime"))
                                                                                BudgetWorkSheetSQL_Value = "'" + Convert.ToDateTime(oRng.get_Value(Type.Missing)).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                                                                            else
                                                                                BudgetWorkSheetSQL_Value = "NULL";
                                                                        else
                                                                            if (oRng.Value2 == null)
                                                                                BudgetWorkSheetSQL_Value = "NULL";
                                                                            else
                                                                            {
                                                                                decimal decimal_out = 0;
                                                                                if (decimal.TryParse(oRng.Value2.ToString().Replace("'", "''"), out decimal_out) == true)
                                                                                    BudgetWorkSheetSQL_Value = oRng.Value2.ToString().Replace("'", "''");
                                                                                else
                                                                                    BudgetWorkSheetSQL_Value = "'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                            }
                                                                    else
                                                                        if ((k.Contains("YrMth") || k.Contains("Date") || k.Contains("date")))
                                                                            if (oRng.get_Value(Type.Missing).GetType() == System.Type.GetType("System.DateTime"))
                                                                                BudgetWorkSheetSQL_Value = BudgetWorkSheetSQL_Value + ",'" + Convert.ToDateTime(oRng.get_Value(Type.Missing)).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                                                                            else
                                                                                BudgetWorkSheetSQL_Value = BudgetWorkSheetSQL_Value + ",NULL";
                                                                        else
                                                                            if (oRng.Value2 == null)
                                                                                BudgetWorkSheetSQL_Value = BudgetWorkSheetSQL_Value + ",NULL";
                                                                            else
                                                                            {
                                                                                decimal decimal_out = 0;
                                                                                if (decimal.TryParse(oRng.Value2.ToString().Replace("'", "''"), out decimal_out) == true)
                                                                                    BudgetWorkSheetSQL_Value = BudgetWorkSheetSQL_Value + "," + oRng.Value2.ToString().Replace("'", "''");
                                                                                else
                                                                                    BudgetWorkSheetSQL_Value = BudgetWorkSheetSQL_Value + ",'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                            }*/

                                                                }
                                                                //BudgetWorkSheetSQL = "insert into Budget_CashFlow_Worksheet (" + BudgetWorkSheetSQL_FieldName + ") values (" + BudgetWorkSheetSQL_Value + ")";
                                                                //ExecuteDatabase(BudgetWorkSheetSQL);

                                                                if (oSheet.Name == Budget_Project_No) // Check Sheet Name = Project No in Excel
                                                                {
                                                                    connection.Open();
                                                                    transa = connection.BeginTransaction();
                                                                    command.Transaction = transa;
                                                                    command.Connection = connection;
                                                                    command.ExecuteNonQuery();
                                                                    transa.Commit();

                                                                    Worksheet_ID = command.Parameters["@Budget_CashFlow_Worksheet_ID"].Value.ToString();
                                                                    string result = command.Parameters["@Result"].Value.ToString();

                                                                    command.Dispose();
                                                                    transa.Dispose();
                                                                    connection.Close();
                                                                }
                                                                once_worksheet = true;
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                sw.WriteLine(oSheet.Name + ": Stored Proc sp_Budget_CashFlow_Worksheet_Update error " + ex.Message.ToString());
                                                                number_of_error = number_of_error + 1;
                                                                connection.Close();
                                                            }
                                                        }
                                                    }

                                                    if (oSheet.Name == Budget_Project_No && number_of_error == 0) // Check Sheet Name = Project No in Excel
                                                    {
                                                        string Budget_CashFlow_Project_ID = "";
                                                        using (SqlConnection connection = new SqlConnection(_connectionString))
                                                        {
                                                            try
                                                            {
                                                                SqlTransaction transa = null;
                                                                SqlCommand command = new SqlCommand("sp_Budget_CashFlow_Project_Update", connection);
                                                                command.CommandType = CommandType.StoredProcedure;

                                                                int id = 0;
                                                                SqlParameter returnValue2 = new SqlParameter("@Budget_CashFlow_Project_ID", id); // ok
                                                                returnValue2.Direction = ParameterDirection.Output;

                                                                //SqlParameter returnValue2 = new SqlParameter("@Budget_CashFlow_Project_ID", id); // ok
                                                                //returnValue2.Direction = ParameterDirection.InputOutput;

                                                                command.Parameters.Add(returnValue2);

                                                                //command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Worksheet_ID", "0"));
                                                                string r = "";

                                                                if (year_end_date.Month < 10)
                                                                    r = year_end_date.Year.ToString() + "0" + year_end_date.Month.ToString();
                                                                else
                                                                    r = year_end_date.Year.ToString() + year_end_date.Month.ToString();

                                                                command.Parameters.Add(new SqlParameter("@YrEnd_CalPeriod", r));
                                                                command.Parameters.Add(new SqlParameter("@Budget_Project_No", Budget_Project_No));
                                                                command.Parameters.Add(new SqlParameter("@Last_Upload_Budget_CashFlow_Worksheet_ID", Worksheet_ID));

                                                                SqlParameter returnValue = new SqlParameter("@Result", ""); // ok
                                                                returnValue.Direction = ParameterDirection.Output;

                                                                command.Parameters.Add(returnValue);


                                                                connection.Open();
                                                                transa = connection.BeginTransaction();
                                                                command.Transaction = transa;
                                                                command.Connection = connection;
                                                                command.ExecuteNonQuery();
                                                                transa.Commit();

                                                                Budget_CashFlow_Project_ID = command.Parameters["@Budget_CashFlow_Project_ID"].Value.ToString();
                                                                string result = command.Parameters["@Result"].Value.ToString();

                                                                command.Dispose();
                                                                transa.Dispose();
                                                                connection.Close();
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                sw.WriteLine(oSheet.Name + ": Stored Proc sp_Budget_CashFlow_Project_Update error " + ex.Message.ToString());
                                                                number_of_error = number_of_error + 1;
                                                                connection.Close();
                                                            }
                                                        }
                                                        string BudgetDetailSQL = "Insert into ";
                                                        string BudgetDetailSQL_FieldName = "";
                                                        string BudgetDetailSQL_Value = "";
                                                        // Insert into

                                                        foreach (string[] k in BudgetDetailDict.Keys)
                                                        {
                                                            //id[0] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                                            //id[1] = dr["Budget_CashFlow_ReportItem_Name"].ToString();
                                                            //value[0] = dr["Row_No"].ToString();
                                                            //value[1] = dr["Budget_CashFlow_ReportItem_ID"].ToString();
                                                            //value[2] = dr["Budget_CashFlow_ReportItem_Code"].ToString();

                                                            string[] row = BudgetDetailDict[k];
                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                BudgetDetailSQL_FieldName = "Budget_CashFlow_Worksheet_ID";
                                                            else
                                                                BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + ",Budget_CashFlow_Worksheet_ID";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                BudgetDetailSQL_FieldName = "Budget_CashFlow_ReportItem_Name";
                                                            else
                                                                BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + ",Budget_CashFlow_ReportItem_Name";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                BudgetDetailSQL_FieldName = "Budget_CashFlow_ReportItem_ID";
                                                            else
                                                                BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + ",Budget_CashFlow_ReportItem_ID";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                BudgetDetailSQL_FieldName = "Budget_CashFlow_ReportItem_Code";
                                                            else
                                                                BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + ",Budget_CashFlow_ReportItem_Code";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                BudgetDetailSQL_FieldName = "Budget_CashFlow_Type";
                                                            else
                                                                BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + ",Budget_CashFlow_Type";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                                BudgetDetailSQL_Value = Worksheet_ID.Replace("'", "''");
                                                            else
                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + "," + Worksheet_ID.Replace("'", "''");

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                                BudgetDetailSQL_Value = "'" + k[1].Replace("'", "''") + "'";
                                                            else
                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + k[1].Replace("'", "''") + "'";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                                BudgetDetailSQL_Value = "'" + row[1].Replace("'", "''") + "'";
                                                            else
                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + row[1].Replace("'", "''") + "'";

                                                            if (Worksheet_ID == "1")
                                                                Console.Write("");
                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                            {
                                                                if (row[2].Replace("'", "''") == "COMPULSORY" || row[2].Replace("'", "''") == "NO" || row[2].Replace("'", "''") == "WARNING")
                                                                {
                                                                    //buster2
                                                                    int count = 0;
                                                                    foreach (string k2 in Budget_CashFlow_ReportItem_CodeDict.Keys)
                                                                    {
                                                                        if (k2 == "Budget_CashFlow_ReportItem_Code")
                                                                        {
                                                                            string[] data = new string[2];
                                                                            Budget_CashFlow_ReportItem_CodeDict.TryGetValue(k2, out data);
                                                                            string col = data[0];

                                                                            if (!string.IsNullOrEmpty(col))
                                                                            {
                                                                                oRng = oSheet.get_Range(col + row[0], col + row[0]);
                                                                                if (Worksheet_ID == "1")
                                                                                    Console.Write("");
                                                                                if (oRng.Value2 != null)
                                                                                    BudgetDetailSQL_Value = "'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                                else
                                                                                    BudgetDetailSQL_Value = "NULL";
                                                                            }
                                                                            else
                                                                                BudgetDetailSQL_Value = "NULL";
                                                                            count = count + 1;
                                                                        }
                                                                    }
                                                                    if (count == 0)
                                                                        BudgetDetailSQL_Value = "NULL";
                                                                }
                                                                else
                                                                    BudgetDetailSQL_Value = "'" + row[2].Replace("'", "''") + "'";
                                                            }
                                                            else
                                                            {
                                                                if (row[2].Replace("'", "''") == "COMPULSORY" || row[2].Replace("'", "''") == "NO" || row[2].Replace("'", "''") == "WARNING")
                                                                {
                                                                    //buster2
                                                                    int count = 0;
                                                                    foreach (string k2 in Budget_CashFlow_ReportItem_CodeDict.Keys)
                                                                    {
                                                                        if (k2 == "Budget_CashFlow_ReportItem_Code")
                                                                        {
                                                                            string[] data = new string[2];
                                                                            Budget_CashFlow_ReportItem_CodeDict.TryGetValue(k2, out data);
                                                                            string col = data[0];

                                                                            if (!string.IsNullOrEmpty(col))
                                                                            {
                                                                                oRng = oSheet.get_Range(col + row[0], col + row[0]);
                                                                                if (Worksheet_ID == "1")
                                                                                    Console.Write("");
                                                                                if (oRng.Value2 != null)
                                                                                    BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                                else
                                                                                    BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",NULL";
                                                                            }
                                                                            else
                                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",NULL";
                                                                            count = count + 1;
                                                                        }
                                                                    }
                                                                    if (count == 0)
                                                                        BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",NULL";
                                                                }
                                                                else
                                                                    BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + row[2].Replace("'", "''") + "'";
                                                            }

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                                BudgetDetailSQL_Value = "'" + row[3].Replace("'", "''") + "'";
                                                            else
                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + row[3].Replace("'", "''") + "'";

                                                            foreach (string k2 in BudgetDetailColDict.Keys)
                                                            {
                                                                string[] data = new string[2];
                                                                BudgetDetailColDict.TryGetValue(k2, out data);
                                                                string col = data[0];

                                                                if (k2 == "ForPeriod21") Console.WriteLine("");
                                                                if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                    BudgetDetailSQL_FieldName = k2;
                                                                else
                                                                    BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + "," + k2;

                                                                if (k2.StartsWith("CumTo_ForPeriod"))
                                                                    Console.WriteLine("");

                                                                if (!string.IsNullOrEmpty(col))
                                                                {
                                                                    oRng = oSheet.get_Range(col + row[0], col + row[0]);

                                                                    if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                                        if (oRng.Value2 == null)
                                                                            BudgetDetailSQL_Value = "NULL";
                                                                        else
                                                                        {
                                                                            if (oRng.Value2.ToString().StartsWith("Gross Value for Own Works"))
                                                                                Console.WriteLine("");
                                                                            if (oRng.Value2 != null)
                                                                                BudgetDetailSQL_Value = "'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                        }
                                                                    else
                                                                        if (oRng.Value2 == null)
                                                                            BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",NULL";
                                                                        else
                                                                        {
                                                                            if (oRng.Value2.ToString().StartsWith("Gross Value for Own Works"))
                                                                                Console.WriteLine("");
                                                                            if (oRng.Value2 != null)
                                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                        }
                                                                }
                                                            }
                                                            BudgetDetailSQL = "Insert into Budget_CashFlow_Detail (" + BudgetDetailSQL_FieldName + ") values (" + BudgetDetailSQL_Value + ")";
                                                            string error = "";
                                                            if (number_of_error == 0)
                                                                error = ExecuteDatabase(BudgetDetailSQL);

                                                            if (!string.IsNullOrEmpty(error))
                                                            {
                                                                sw.WriteLine(oSheet.Name + ": Insert into Budget_CashFlow_Detail error " + error);
                                                                number_of_error = number_of_error + 1;
                                                            }
                                                            BudgetDetailSQL = "";
                                                            BudgetDetailSQL_FieldName = "";
                                                            BudgetDetailSQL_Value = "";
                                                        }
                                                        //Console.WriteLine("");
                                                        using (SqlConnection connection = new SqlConnection(_connectionString))
                                                        {
                                                            try
                                                            {
                                                                SqlTransaction transa = null;
                                                                SqlCommand command = new SqlCommand("sp_Update_Worksheet_Status", connection);
                                                                command.CommandType = CommandType.StoredProcedure;

                                                                command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Worksheet_ID", Worksheet_ID)); // ok
                                                                if (number_of_error == 0)
                                                                    command.Parameters.Add(new SqlParameter("@Upload_Status", "SUCCESS"));//ok
                                                                else
                                                                    command.Parameters.Add(new SqlParameter("@Upload_Status", "FAILED"));//ok

                                                                SqlParameter returnValue = new SqlParameter("@Result", ""); // ok
                                                                returnValue.Direction = ParameterDirection.Output;

                                                                command.Parameters.Add(returnValue);


                                                                connection.Open();
                                                                transa = connection.BeginTransaction();
                                                                command.Transaction = transa;
                                                                command.Connection = connection;
                                                                command.ExecuteNonQuery();
                                                                transa.Commit();

                                                                //Budget_CashFlow_Project_ID = command.Parameters["@Budget_CashFlow_Worksheet_ID"].Value.ToString();
                                                                string result = command.Parameters["@Result"].Value.ToString();

                                                                command.Dispose();
                                                                transa.Dispose();
                                                                connection.Close();
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                sw.WriteLine(oSheet.Name + ": Update database status error " + ex.Message.ToString());
                                                                connection.Close();
                                                            }
                                                        }
                                                        if (number_of_error > 0)
                                                            error_sheet.Add(error_sheet.Count.ToString(), oSheet.Name);
                                                    }
                                                }
                                            }
                                        }
                                        string errorMsg = "";
                                        foreach (string k2 in error_sheet.Keys)
                                        {
                                            string sheetname = "";
                                            error_sheet.TryGetValue(k2, out sheetname);
                                            if (string.IsNullOrEmpty(errorMsg))
                                                errorMsg = errorMsg + sheetname;
                                            else
                                                errorMsg = errorMsg + "," + sheetname;
                                        }
                                        if (error_sheet.Count == 0)
                                        {
                                            upload_status.Add(oSheet.Name, "SUCCESS");
                                            error_status.Add(oSheet.Name, "");

                                            if (isWrittenEndHtml == false)
                                            {
                                                //sw2.WriteLine("     <td>");
                                                //sw2.WriteLine("         SUCCESS");
                                                //sw2.WriteLine("     </td>");
                                                //sw2.WriteLine("     <td>");
                                                //sw2.WriteLine("         .");
                                                //sw2.WriteLine("     </td>");

                                            }
                                            //MessageBox.Show("Upload Success");
                                        }
                                        else
                                        {
                                            upload_status.Add(oSheet.Name, "FAILED");
                                            error_status.Add(oSheet.Name, "         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");

                                            if (isWrittenEndHtml == false)
                                            {
                                                //sw2.WriteLine("     <td>");
                                                //sw2.WriteLine("         FAILED");
                                                //sw2.WriteLine("     </td>");
                                                //sw2.WriteLine("     <td>");
                                                //sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                                //sw2.WriteLine("     </td>");

                                            }
                                            //MessageBox.Show("upload to database Error worksheet " + errorMsg);
                                        }
                                    }

                                    /*
                                    string errorMsg = "";
                                    foreach (string k2 in error_sheet.Keys)
                                    {
                                        string sheetname = "";
                                        error_sheet.TryGetValue(k2, out sheetname);
                                        if (string.IsNullOrEmpty(errorMsg))
                                            errorMsg = errorMsg + sheetname;
                                        else
                                            errorMsg = errorMsg + "," + sheetname;
                                    }
                                    if (error_sheet.Count == 0)
                                    {
                                        upload_status.Add("SUCCESS");
                                        error_status.Add("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");

                                        if (isWrittenEndHtml == false)
                                        {
                                            
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         SUCCESS");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            //sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            sw2.WriteLine("         .");
                                            sw2.WriteLine("     </td>");
                                             
                                        }
                                        //MessageBox.Show("Upload Success");
                                    }
                                    else
                                    {
                                        upload_status.Add("FAILED");
                                        error_status.Add("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");

                                        if (isWrittenEndHtml == false)
                                        {
                                            
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         FAILED");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            sw2.WriteLine("     </td>");
                                            
                                        }
                                        MessageBox.Show("upload to database Error worksheet " + errorMsg);
                                    }*/
                                }
                                else
                                {
                                    /*
                                    if (error_sheet.Count == 0)
                                    {
                                        upload_status.Add("SUCCESS");
                                        error_status.Add("");

                                        if (isWrittenEndHtml == false)
                                        {
                                            
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         SUCCESS");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            //sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            sw2.WriteLine("         .");
                                            sw2.WriteLine("     </td>");
                                             
                                        }
                                        //MessageBox.Show("Upload Success");
                                    }
                                    else
                                    {
                                        upload_status.Add("FAILED");
                                        error_status.Add("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                        if (isWrittenEndHtml == false)
                                        {
                                            
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         FAILED");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            sw2.WriteLine("     </td>");
                                            
                                        }
                                        //MessageBox.Show("upload to database Error worksheet " + errorMsg);
                                    }*/
                                }
                            }
                            /*
                            if (oSheet.Cells[3, 1] == "Project Name")
                            {
                                oRng = oSheet.get_Range("A1", "A65536");
                                oRng.ColumnWidth = 114;
                            }

                            oRng = oSheet.get_Range("A1", "A65536");
                            //oXL.Visible = true;
                            */
                            /*
                            Boolean one_of_failed = false;
                            foreach (String k in log_sheet_name.Keys)
                            {
                                String dummy = "";
                                if(upload_status.TryGetValue(k, out dummy) == false)
                                {
                                    upload_status.Add(k, "FAILED");
                                }

                                dummy = "";
                                if(error_status.TryGetValue(k, out dummy)== false)
                                {
                                    error_status.Add(k, "");
                                }                                
                            }
                            foreach (String k in log_sheet_name.Keys)
                            {
                                String dummy = "";
                                sw2.WriteLine("<tr width=70>");
                                sw2.WriteLine("     <td>");
                                log_full_name.TryGetValue(k, out dummy);
                                sw2.WriteLine(dummy);
                                sw2.WriteLine("     </td>");
                                sw2.WriteLine("     <td>");
                                dummy = "";
                                log_sheet_name.TryGetValue(k, out dummy);
                                sw2.WriteLine(dummy);
                                sw2.WriteLine("     </td>");
                                sw2.WriteLine("     <td>");
                                dummy = "";
                                validation_status.TryGetValue(k, out dummy);
                                sw2.WriteLine(dummy);
                                sw2.WriteLine("     </td>");
                                sw2.WriteLine("     <td>");
                                dummy = "";
                                upload_status.TryGetValue(k, out dummy);
                                sw2.WriteLine(dummy);
                                sw2.WriteLine("     </td>");
                                sw2.WriteLine("     <td>");
                                dummy = "";
                                error_status.TryGetValue(k, out dummy);
                                sw2.WriteLine(dummy);
                                sw2.WriteLine("     </td>");
                                sw2.WriteLine("</tr>");
                            }
                            */
                            oWB.Close(null, null, null);
                            oXL.Workbooks.Close();
                            oXL.Quit();
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(oRng);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB);
                            oSheet = null;
                            oWB = null;
                            oXL = null;
                            GC.Collect();

                            //Application.Exit();
                        }
                        catch (Exception ex)
                        {
                            //Error_Label.Text = "No file generated! " + " Error Occur - Generation of Excel " + ex.Message;
                            oWB.Close(null, null, null);
                            oXL.Workbooks.Close();
                            oXL.Quit();
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(oRng);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB);
                            oSheet = null;
                            oWB = null;
                            oXL = null;
                            GC.Collect();
                        }
                        #endregion one
                    }
                    //many
                    if (param[0].ToUpper() == "MANY")
                    {
                        #region many
                        object Missing = System.Type.Missing;
                        Excel._Workbook oWB = null;
                        Excel._Worksheet oSheet = null;
                        Excel.Range oRng = null;
                        Excel.Application oXL = new Excel.Application();
                        try
                        {
                            oXL.Visible = false;
                            oXL.DisplayAlerts = false;

                            //Get a new workbook.
                            //oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing));

                            oWB = oXL.Workbooks.Open(fi.FullName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing);

                            string CashFlow_Header_ID = "";
                            bool once_header = false;
                            string Budget_Project_No = "";
                            bool once_worksheet = false;

                            // validation

                            string error_filename = "error_" + fi.Name + "_" + DateTime.Now.ToString().Replace(":", "-").Replace("/", "-");
                            string error_path = param[2] + "\\" + error_filename + ".txt";
                            FileStream file = new FileStream(error_path, FileMode.Append);
                            //FileStream file = new FileStream(Environment.CurrentDirectory + "\\error.txt", FileMode.Append);
                            bool isWrittenEndHtml = false;
                            using (StreamWriter sw = new StreamWriter(file))
                            {

                                for (int i = 1; i <= oWB.Sheets.Count; i++)
                                {
                                    //oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                                    oSheet = (Excel._Worksheet)oWB.Sheets[i];

                                    Boolean valid = true;
                                    // Validate Header
                                    string[] types = new string[2];
                                    types[0] = "BUDGET";
                                    types[1] = "CASHFLOW";
                                    once_worksheet = false;
                                    string Worksheet_ID = "";
                                    int number_of_error = 0;
                                    bool user_custom_sheet = false;
                                    int counter3 = 0;
                                    foreach (string type in types)
                                    {
                                        counter3 = counter3 + 1;
                                        string latestperiod = LatestPeriodInReport();
                                        //string startrow = StartRow(type, latestperiod);

                                        //select * from dbo.Budget_CashFlow_Excel_ReportItem_Detail where Heading_Type = 'D'
                                        Dictionary<string, string[]> BudgetHeaderDict = ReadHeaderFromDB(type);
                                        foreach (string k in BudgetHeaderDict.Keys)
                                        {
                                            string[] location = BudgetHeaderDict[k];
                                            oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);

                                        }

                                        oRng = oSheet.get_Range("A1", "A1");
                                        if (oRng.Value2 == null)
                                        {
                                            user_custom_sheet = true;
                                            number_of_error = number_of_error + 1;
                                        }
                                        else
                                            if (string.IsNullOrEmpty(oRng.Value2.ToString()))
                                                number_of_error = number_of_error + 1;

                                        if (oRng.Value2 != null)
                                        {
                                            if (!string.IsNullOrEmpty(oRng.Value2.ToString()))
                                            {
                                                if (counter3 == 1)
                                                {
                                                    try
                                                    {
                                                        log_full_name.Add(oSheet.Name, fi.FullName);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show("Same project code among files");
                                                    }
                                                    log_sheet_name.Add(oSheet.Name, oSheet.Name);

                                                    //sw2.WriteLine(" <tr>");
                                                    //sw2.WriteLine("     <td width=50>");
                                                    //sw2.WriteLine("         " + fi.FullName);
                                                    //sw2.WriteLine("     </td>");
                                                    //sw2.WriteLine("     <td>");
                                                    //sw2.WriteLine("         " + oSheet.Name);
                                                    //sw2.WriteLine("     </td>");
                                                }
                                                int year_end_index = oRng.Value2.ToString().IndexOf("The Year Ended");
                                                //string year_end_date_string = oRng.Value2.ToString().Substring(year_end_index, oRng.Value2.ToString().Length - year_end_index).Replace("The Year Ended", "").Trim();
                                                //string year_end_date_string = oRng.Value2.ToString().Substring(year_end_index, oRng.Value2.ToString().Length - year_end_index).Replace("The Year Ended", "").Trim();
                                                //DateTime year_end_date = Convert.ToDateTime(year_end_date_string);

                                                DateTime year_end_date = Convert.ToDateTime("1-" + (Convert.ToInt32(param[4].Substring(4, 2)) + 1).ToString() + "-" + param[4].Substring(0, 4)).AddDays(-1);

                                                string BudgetWorkSheetSQL = "Insert into ";
                                                string BudgetWorkSheetSQL_FieldName = "";
                                                string BudgetWorkSheetSQL_Value = "";

                                                string[] Budget_Project_On_Hand_Status_location = ReadBudget_Project_On_Hand_StatusFromDB(latestperiod);
                                                oRng = oSheet.get_Range(Budget_Project_On_Hand_Status_location[1] + Budget_Project_On_Hand_Status_location[0], Budget_Project_On_Hand_Status_location[1] + Budget_Project_On_Hand_Status_location[0]);
                                                string Budget_Project_On_Hand_Status = oRng.Value2.ToString();


                                                Dictionary<string, string[]> BudgetHeaderItemRowColDict = ReadHeaderItemRowColFromDB(Budget_Project_On_Hand_Status, latestperiod);
                                                Dictionary<string[], string[]> BudgetDetailDict = ReadDetailFromDB(type, Budget_Project_On_Hand_Status);
                                                Dictionary<string, string[]> BudgetDetailColDict = ReadDetailColFromDB(type);
                                                Dictionary<string, string[]> Budget_CashFlow_ReportItem_CodeDict = ReadBudget_CashFlow_ReportItem_CodeFromDB(type);
                                                Dictionary<string, string[]> Budget_CashFlow_AllReportItem_CodeDict = ReadBudget_CashFlow_AllReportItem_CodeFromDB(type);

                                                string _connectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;

                                                if (once_worksheet == false)
                                                {
                                                    using (SqlConnection connection = new SqlConnection(_connectionString))
                                                    {
                                                        try
                                                        {
                                                            int id = 0;

                                                            foreach (string k in BudgetHeaderItemRowColDict.Keys)
                                                            {
                                                                string[] location = new string[2];
                                                                BudgetHeaderItemRowColDict.TryGetValue(k, out location);
                                                                oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);


                                                                if ((k.Contains("YrMth") || k.Contains("Date") || k.Contains("date")))
                                                                {
                                                                    /*
                                                                    DateTime datetime_parse_result;
                                                                    if (DateTime.TryParse(oRng.get_Value(Type.Missing).ToString().Replace("After", "").Replace(" ", "").Trim(), out datetime_parse_result) == true)
                                                                    {
                                                                        //command.Parameters.Add(new SqlParameter("@" + k, Convert.ToDateTime(oRng.get_Value(Type.Missing)).ToString("yyyy/MM/dd HH:mm:ss")));
                                                                        Console.Write("");
                                                                    }
                                                                    else
                                                                    {
                                                                        number_of_error = number_of_error + 1;
                                                                        sw.WriteLine(oSheet.Name + ": " + k + " is not a date");
                                                                    }*/
                                                                }
                                                                else
                                                                    if (oRng.Value2 == null)
                                                                    {
                                                                        Console.Write("");
                                                                    }
                                                                    else
                                                                    {
                                                                        decimal decimal_out = 0;
                                                                        if (decimal.TryParse(oRng.Value2.ToString().Replace("'", "''"), out decimal_out) == true)
                                                                        {
                                                                            if (k.Contains("Budget_Project_No"))
                                                                                Budget_Project_No = oRng.Value2.ToString().Replace("'", "''");
                                                                        }
                                                                        else
                                                                            if (k.Contains("Hit_Rate"))
                                                                            {
                                                                                Console.Write("");
                                                                                //command.Parameters.Add(new SqlParameter("@" + k, decimal_out));
                                                                            }
                                                                            else
                                                                            {
                                                                                if (k.Contains("Budget_Project_No"))
                                                                                    Budget_Project_No = oRng.Value2.ToString().Replace("'", "''");
                                                                            }
                                                                    }
                                                            }
                                                            //BudgetWorkSheetSQL = "insert into Budget_CashFlow_Worksheet (" + BudgetWorkSheetSQL_FieldName + ") values (" + BudgetWorkSheetSQL_Value + ")";
                                                            //ExecuteDatabase(BudgetWorkSheetSQL);

                                                            once_worksheet = true;
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            connection.Close();
                                                        }
                                                    }
                                                    string number_of_found = ProjectCodeExistInPCMS(Budget_Project_No);
                                                    if (Budget_Project_On_Hand_Status == "Unsecured")
                                                    {
                                                        if (Budget_Project_No.Contains(" "))
                                                        {
                                                            number_of_error = number_of_error + 1;
                                                            sw.WriteLine(oSheet.Name + ": Budget_Project_No Contain space");
                                                        }
                                                        if (Budget_Project_No.Length > 8)
                                                        {
                                                            number_of_error = number_of_error + 1;
                                                            sw.WriteLine(oSheet.Name + ": Budget_Project_No length > 8");
                                                        }
                                                    }

                                                    if (string.IsNullOrEmpty(number_of_found) && Budget_Project_On_Hand_Status != "Unsecured")
                                                    {
                                                        number_of_error = number_of_error + 1;
                                                        sw.WriteLine(oSheet.Name + ": Project Code/No not exist in PCMS");
                                                    }
                                                    else
                                                    {
                                                        int num = 0;
                                                        if (Int32.TryParse(number_of_found, out num) == false && Budget_Project_On_Hand_Status != "Unsecured")
                                                        {
                                                            number_of_error = number_of_error + 1;
                                                            sw.WriteLine(oSheet.Name + ": Project Code/No not exist in PCMS");
                                                        }
                                                        if (num == 0 && Budget_Project_On_Hand_Status != "Unsecured")
                                                        {
                                                            number_of_error = number_of_error + 1;
                                                            sw.WriteLine(oSheet.Name + ": Project Code/No not exist in PCMS");
                                                        }
                                                    }
                                                    //if (oSheet.Name != Budget_Project_No)
                                                    //sw.WriteLine(k[1] + " Missing Report Code");
                                                }
                                                string Budget_CashFlow_Project_ID = "";
                                                string BudgetDetailSQL = "Insert into ";
                                                string BudgetDetailSQL_FieldName = "";
                                                string BudgetDetailSQL_Value = "";
                                                // Insert into
                                                string ErrorReport = "";
                                                if (!string.IsNullOrEmpty(Budget_Project_No))
                                                {
                                                    foreach (string[] k in BudgetDetailDict.Keys)
                                                    {
                                                        //id[0] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                                        //id[1] = dr["Budget_CashFlow_ReportItem_Name"].ToString();
                                                        //value[0] = dr["Row_No"].ToString();
                                                        //value[1] = dr["Budget_CashFlow_ReportItem_ID"].ToString();
                                                        //value[2] = dr["Budget_CashFlow_ReportItem_Code"].ToString();

                                                        string[] row = BudgetDetailDict[k];

                                                        if (Worksheet_ID == "1")
                                                            Console.Write("");
                                                        if (row[2].Replace("'", "''") == "COMPULSORY" || row[2].Replace("'", "''") == "WARNING")
                                                        {
                                                            foreach (string k2 in Budget_CashFlow_ReportItem_CodeDict.Keys)
                                                            {
                                                                if (k2 == "Budget_CashFlow_ReportItem_Code") // i.e. get column C
                                                                {
                                                                    string[] data = new string[2];
                                                                    Budget_CashFlow_ReportItem_CodeDict.TryGetValue(k2, out data);
                                                                    string col = data[0];

                                                                    if (!string.IsNullOrEmpty(data[1]))
                                                                    {
                                                                        switch (data[1])
                                                                        {
                                                                            /*
                                                                        case "C":
                                                                            DateTime datetime_result2;
                                                                            if (DateTime.TryParse(oRng.Value2.ToString(), out datetime_result2) == true)
                                                                            {
                                                                                number_of_error = number_of_error + 1;
                                                                                sw.WriteLine(oSheet.Name + ":It is not a String - " + k[1]);
                                                                            }
                                                                            decimal numeric_result2 = 0;
                                                                            if (Decimal.TryParse(oRng.Value2.ToString(), out numeric_result2) == true)
                                                                            {
                                                                                number_of_error = number_of_error + 1;
                                                                                sw.WriteLine(oSheet.Name + ":It is not a String - " + k[1]);
                                                                            }
                                                                            break;*/
                                                                            case "D":
                                                                                DateTime datetime_result;
                                                                                if (DateTime.TryParse(oRng.Value2.ToString(), out datetime_result) == false)
                                                                                {
                                                                                    number_of_error = number_of_error + 1;
                                                                                    sw.WriteLine(oSheet.Name + ":It is not a datetime format - " + k[1]);
                                                                                }
                                                                                break;
                                                                            case "N":
                                                                                decimal numeric_result = 0;
                                                                                if (Decimal.TryParse(oRng.Value2.ToString(), out numeric_result) == false)
                                                                                {
                                                                                    number_of_error = number_of_error + 1;
                                                                                    sw.WriteLine(oSheet.Name + ":It is not numeric - " + k[1]);
                                                                                }
                                                                                break;
                                                                            default:
                                                                                break;
                                                                        }
                                                                    }
                                                                    if (!string.IsNullOrEmpty(col))
                                                                    {
                                                                        oRng = oSheet.get_Range(col + row[0], col + row[0]);
                                                                        if (Worksheet_ID == "1")
                                                                            Console.Write("");
                                                                        if (oRng.Value2 != null)
                                                                        {
                                                                            if (type == "BUDGET")
                                                                            {
                                                                                string number_of_found = CostCodeExistInPCMS(oRng.Value2.ToString().Replace("'", "''"));
                                                                                if (string.IsNullOrEmpty(number_of_found))
                                                                                {
                                                                                    number_of_error = number_of_error + 1;
                                                                                    sw.WriteLine(oSheet.Name + ":Missing Report Code " + col + row[0] + " - " + k[1]);
                                                                                }
                                                                                else
                                                                                {
                                                                                    int num = 0;
                                                                                    if (Int32.TryParse(number_of_found, out num) == false)
                                                                                    {
                                                                                        number_of_error = number_of_error + 1;
                                                                                        sw.WriteLine(oSheet.Name + ":Missing CashFlow Code and number of found is not an integer " + col + row[0] + " - " + k[1]);
                                                                                    }
                                                                                    if (num == 0)
                                                                                    {
                                                                                        number_of_error = number_of_error + 1;
                                                                                        sw.WriteLine(oSheet.Name + ":Missing Report Code " + col + row[0] + " - " + k[1]);
                                                                                    }
                                                                                }
                                                                            }
                                                                            else if (type == "CASHFLOW")
                                                                            {
                                                                                string number_of_found = CashFlowCodeExistInPCMS(oRng.Value2.ToString().Replace("'", "''"));
                                                                                if (string.IsNullOrEmpty(number_of_found))
                                                                                {
                                                                                    number_of_error = number_of_error + 1;
                                                                                    sw.WriteLine(oSheet.Name + ":Missing CashFlow Code " + col + row[0] + " - " + k[1]);
                                                                                }
                                                                                else
                                                                                {
                                                                                    int num = 0;
                                                                                    if (Int32.TryParse(number_of_found, out num) == false)
                                                                                    {
                                                                                        number_of_error = number_of_error + 1;
                                                                                        sw.WriteLine(oSheet.Name + ":Missing CashFlow Code and number of found is not an integer " + col + row[0] + " - " + k[1]);
                                                                                    }
                                                                                    if (num == 0)
                                                                                    {
                                                                                        number_of_error = number_of_error + 1;
                                                                                        sw.WriteLine(oSheet.Name + ":Missing CashFlow Code " + col + row[0] + " - " + k[1]);
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            if ((row[2].Replace("'", "''") == "COMPULSORY" || row[2].Replace("'", "''") == "WARNING"))
                                                                            {
                                                                                if (type == "BUDGET")
                                                                                {
                                                                                    number_of_error = number_of_error + 1;
                                                                                    sw.WriteLine(oSheet.Name + ":Missing Report Code " + col + row[0] + " - " + k[1]);
                                                                                }
                                                                                else if (type == "CASHFLOW")
                                                                                {
                                                                                    number_of_error = number_of_error + 1;
                                                                                    sw.WriteLine(oSheet.Name + ":Missing CashFlow Code " + col + row[0] + " - " + k[1]);
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else
                                                            BudgetDetailSQL_Value = "'" + row[2].Replace("'", "''") + "'";


                                                        if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                            BudgetDetailSQL_Value = "'" + row[3].Replace("'", "''") + "'";
                                                        else
                                                            BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + row[3].Replace("'", "''") + "'";



                                                        foreach (string k2 in BudgetDetailColDict.Keys)
                                                        {
                                                            string[] data = new string[2];
                                                            BudgetDetailColDict.TryGetValue(k2, out data);
                                                            string col = data[0];

                                                            if (k2 == "ForPeriod21") Console.WriteLine("");
                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                BudgetDetailSQL_FieldName = k2;
                                                            else
                                                                BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + "," + k2;

                                                            if (k2.StartsWith("CumTo_ForPeriod"))
                                                                Console.WriteLine("");

                                                            if (!string.IsNullOrEmpty(col))
                                                            {
                                                                oRng = oSheet.get_Range(col + row[0], col + row[0]);

                                                                if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                                    if (oRng.Value2 == null)
                                                                        BudgetDetailSQL_Value = "NULL";
                                                                    else
                                                                    {
                                                                        if (oRng.Value2.ToString().StartsWith("Gross Value for Own Works"))
                                                                            Console.WriteLine("");
                                                                        if (oRng.Value2 != null)
                                                                            BudgetDetailSQL_Value = "'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                    }
                                                                else
                                                                    if (oRng.Value2 == null)
                                                                        BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",NULL";
                                                                    else
                                                                    {
                                                                        if (oRng.Value2.ToString().StartsWith("Gross Value for Own Works"))
                                                                            Console.WriteLine("");
                                                                        if (oRng.Value2 != null)
                                                                            BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                    }
                                                            }
                                                        }
                                                    }
                                                    //validate header
                                                    foreach (string k3 in BudgetHeaderItemRowColDict.Keys)
                                                    {
                                                        string[] location = new string[3];
                                                        BudgetHeaderItemRowColDict.TryGetValue(k3, out location);
                                                        oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);

                                                        switch (location[2])
                                                        {
                                                            case "D":
                                                                DateTime datetime_result;
                                                                if (!oRng.get_Value(Missing).ToString().Replace("After", "").Contains("N/A"))
                                                                {
                                                                    if (DateTime.TryParse(oRng.get_Value(Missing).ToString().Replace("After", ""), out datetime_result) == false)
                                                                    {
                                                                        number_of_error = number_of_error + 1;
                                                                        sw.WriteLine(oSheet.Name + ":It is not a datetime format at " + location[1] + location[0] + " - " + k3);
                                                                    }
                                                                }
                                                                break;
                                                            case "N":
                                                                decimal numeric_result = 0;
                                                                if (Decimal.TryParse(oRng.get_Value(Missing).ToString(), out numeric_result) == false)
                                                                {
                                                                    if (!oRng.get_Value(Missing).ToString().Replace("After", "").Contains("N/A"))
                                                                    {
                                                                        number_of_error = number_of_error + 1;
                                                                        sw.WriteLine(oSheet.Name + ":It is not numeric at " + location[1] + location[0] + " - " + k3);
                                                                    }
                                                                }
                                                                break;
                                                            default:
                                                                break;
                                                        }
                                                    }
                                                    // validate data
                                                    int counter8 = 0;
                                                    foreach (string[] k in BudgetDetailDict.Keys)
                                                    {
                                                        //id[0] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                                        //id[1] = dr["Budget_CashFlow_ReportItem_Name"].ToString();
                                                        //value[0] = dr["Row_No"].ToString();
                                                        //value[1] = dr["Budget_CashFlow_ReportItem_ID"].ToString();
                                                        //value[2] = dr["Budget_CashFlow_ReportItem_Code"].ToString();

                                                        string[] row = BudgetDetailDict[k];

                                                        foreach (string k2 in Budget_CashFlow_AllReportItem_CodeDict.Keys)
                                                        {
                                                            string[] data = new string[2];
                                                            Budget_CashFlow_AllReportItem_CodeDict.TryGetValue(k2, out data);

                                                            string col = data[0];

                                                            if (!string.IsNullOrEmpty(col))
                                                            {
                                                                oRng = oSheet.get_Range(col + row[0], col + row[0]);
                                                                switch (data[1])
                                                                {
                                                                    case "D":
                                                                        DateTime datetime_result;
                                                                        if (!oRng.get_Value(Missing).ToString().Replace("After", "").Contains("N/A"))
                                                                        {
                                                                            if (oRng.get_Value(Missing) != null)
                                                                            {
                                                                                if (DateTime.TryParse(oRng.get_Value(Missing).ToString().Replace("After", ""), out datetime_result) == false)
                                                                                {
                                                                                    number_of_error = number_of_error + 1;
                                                                                    sw.WriteLine(oSheet.Name + ":It is not a datetime format " + col + row[0] + " - " + k[1]);
                                                                                }
                                                                            }
                                                                        }
                                                                        break;
                                                                    case "N":
                                                                        decimal numeric_result = 0;
                                                                        if (oRng.get_Value(Missing) != null)
                                                                        {
                                                                            if (Decimal.TryParse(oRng.get_Value(Missing).ToString(), out numeric_result) == false)
                                                                            {
                                                                                number_of_error = number_of_error + 1;
                                                                                sw.WriteLine(oSheet.Name + ":It is not numeric " + col + row[0] + " - " + k[1]);
                                                                            }
                                                                        }
                                                                        break;
                                                                    default:
                                                                        break;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (user_custom_sheet == false)
                                    {
                                        if (number_of_error > 0)
                                        {
                                            isWrittenEndHtml = true;

                                            validation_status.Add(oSheet.Name, "FAILED");
                                            upload_status.Add(oSheet.Name, "FAILED");
                                            error_status.Add(oSheet.Name, "         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            /*
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         FAILED");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         FAILED");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine(" </tr>");
                                            */
                                            error_sheet.Add(error_sheet.Count.ToString(), oSheet.Name);
                                        }
                                        else
                                        {
                                            //isWrittenEndHtml = true;
                                            validation_status.Add(oSheet.Name, "SUCCESS");
                                            /*
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         SUCCESS");
                                            sw2.WriteLine("     </td>");
                                            */
                                            //sw2.WriteLine("     <td>");
                                            //sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            //sw2.WriteLine("     </td>");
                                            //sw2.WriteLine(" </tr>");
                                        }
                                    }
                                }
                                //sw.WriteLine("error_sheet.Count:" + error_sheet.Count.ToString());
                                //sw.WriteLine("Budget_Project_No:" + Budget_Project_No);
                                if (error_sheet.Count == 0 && !string.IsNullOrEmpty(Budget_Project_No))
                                {
                                    // process
                                    for (int i = 1; i <= oWB.Sheets.Count; i++)
                                    {
                                        //oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                                        oSheet = (Excel._Worksheet)oWB.Sheets[i];

                                        Boolean valid = true;
                                        // Validate Header
                                        string[] types = new string[2];
                                        types[0] = "BUDGET";
                                        types[1] = "CASHFLOW";
                                        once_worksheet = false;
                                        string Worksheet_ID = "";



                                        foreach (string type in types)
                                        {
                                            string latestperiod = LatestPeriodInReport();
                                            //string startrow = StartRow(type, latestperiod);

                                            //select * from dbo.Budget_CashFlow_Excel_ReportItem_Detail where Heading_Type = 'D'
                                            Dictionary<string, string[]> BudgetHeaderDict = ReadHeaderFromDB(type);
                                            foreach (string k in BudgetHeaderDict.Keys)
                                            {
                                                string[] location = BudgetHeaderDict[k];
                                                oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);

                                            }
                                            /*
                                            startrow = StartRow("CASHFLOW", LatestPeriodInReport());
                    
                                            Dictionary<string, string[]> CashFlowHeaderDict = ReadHeaderFromDB("CASHFLOW");
                                            foreach (string k in CashFlowHeaderDict.Keys)
                                            {
                                                string[] location = BudgetHeaderDict[k];
                                                oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);
                                            }*/
                                            int number_of_error = 0;
                                            oRng = oSheet.get_Range("A1", "A1");
                                            if (oRng.Value2 == null)
                                                number_of_error = number_of_error + 1;
                                            else
                                                if (string.IsNullOrEmpty(oRng.Value2.ToString()))
                                                    number_of_error = number_of_error + 1;

                                            if (oRng.Value2 != null)
                                            {
                                                if (!string.IsNullOrEmpty(oRng.Value2.ToString()))
                                                {

                                                    int year_end_index = oRng.Value2.ToString().IndexOf("The Year Ended");
                                                    //string year_end_date_string = oRng.Value2.ToString().Substring(year_end_index, oRng.Value2.ToString().Length - year_end_index).Replace("The Year Ended", "").Trim();
                                                    //DateTime year_end_date = Convert.ToDateTime(year_end_date_string);
                                                    DateTime year_end_date = Convert.ToDateTime("1-" + (Convert.ToInt32(param[4].Substring(4, 2)) + 1).ToString() + "-" + param[4].Substring(0, 4)).AddDays(-1);

                                                    string BudgetWorkSheetSQL = "Insert into ";
                                                    string BudgetWorkSheetSQL_FieldName = "";
                                                    string BudgetWorkSheetSQL_Value = "";

                                                    string[] Budget_Project_On_Hand_Status_location = ReadBudget_Project_On_Hand_StatusFromDB(latestperiod);
                                                    oRng = oSheet.get_Range(Budget_Project_On_Hand_Status_location[1] + Budget_Project_On_Hand_Status_location[0], Budget_Project_On_Hand_Status_location[1] + Budget_Project_On_Hand_Status_location[0]);
                                                    string Budget_Project_On_Hand_Status = oRng.Value2.ToString();


                                                    Dictionary<string, string[]> BudgetHeaderItemRowColDict = ReadHeaderItemRowColFromDB(Budget_Project_On_Hand_Status, latestperiod);
                                                    Dictionary<string[], string[]> BudgetDetailDict = ReadDetailFromDB(type, Budget_Project_On_Hand_Status);
                                                    Dictionary<string, string[]> BudgetDetailColDict = ReadDetailColFromDB(type);
                                                    Dictionary<string, string[]> Budget_CashFlow_ReportItem_CodeDict = ReadBudget_CashFlow_ReportItem_CodeFromDB(type);

                                                    string _connectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
                                                    if (once_header == false)
                                                    {
                                                        using (SqlConnection connection = new SqlConnection(_connectionString))
                                                        {
                                                            try
                                                            {
                                                                SqlTransaction transa = null;
                                                                SqlCommand command = new SqlCommand("sp_Budget_CashFlow_Header_Update", connection);
                                                                command.CommandType = CommandType.StoredProcedure;

                                                                int id = 0;
                                                                SqlParameter returnValue2 = new SqlParameter("@Budget_CashFlow_Header_ID", id); // ok
                                                                returnValue2.Direction = ParameterDirection.InputOutput;

                                                                command.Parameters.Add(returnValue2);

                                                                command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Upload_By", param[1]));

                                                                SqlParameter returnValue = new SqlParameter("@Result", ""); // ok
                                                                returnValue.Direction = ParameterDirection.Output;

                                                                command.Parameters.Add(returnValue);

                                                                connection.Open();
                                                                transa = connection.BeginTransaction();
                                                                command.Transaction = transa;
                                                                command.Connection = connection;
                                                                command.ExecuteNonQuery();
                                                                transa.Commit();

                                                                CashFlow_Header_ID = command.Parameters["@Budget_CashFlow_Header_ID"].Value.ToString();
                                                                string result = command.Parameters["@Result"].Value.ToString();

                                                                command.Dispose();
                                                                transa.Dispose();
                                                                connection.Close();
                                                                once_header = true;
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                sw.WriteLine(oSheet.Name + ": Stored Proc sp_Budget_CashFlow_Header_Update error " + ex.Message.ToString());
                                                                number_of_error = number_of_error + 1;
                                                                connection.Close();
                                                            }
                                                        }
                                                    }

                                                    if (once_worksheet == false && number_of_error == 0)
                                                    {
                                                        using (SqlConnection connection = new SqlConnection(_connectionString))
                                                        {
                                                            try
                                                            {
                                                                SqlTransaction transa = null;
                                                                SqlCommand command = new SqlCommand("sp_Budget_CashFlow_Worksheet_Update", connection);
                                                                command.CommandType = CommandType.StoredProcedure;

                                                                int id = 0;
                                                                SqlParameter returnValue2 = new SqlParameter("@Budget_CashFlow_Worksheet_ID", id); // ok
                                                                returnValue2.Direction = ParameterDirection.InputOutput;

                                                                command.Parameters.Add(returnValue2);

                                                                //command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Worksheet_ID", "0"));
                                                                command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Header_ID", CashFlow_Header_ID));
                                                                command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Worksheet_Name", oSheet.Name));
                                                                command.Parameters.Add(new SqlParameter("@YrEnd_Date", year_end_date.ToString("yyyy/MM/dd HH:mm:ss")));
                                                                command.Parameters.Add(new SqlParameter("@Budget_Project_On_Hand_Status", Budget_Project_On_Hand_Status));
                                                                command.Parameters.Add(new SqlParameter("@Upload_FileName", fi.FullName));

                                                                SqlParameter returnValue = new SqlParameter("@Result", ""); // ok
                                                                returnValue.Direction = ParameterDirection.Output;

                                                                command.Parameters.Add(returnValue);

                                                                //command.Parameters.Add(new SqlParameter("@Result", ""));


                                                                foreach (string k in BudgetHeaderItemRowColDict.Keys)
                                                                {
                                                                    /*
                                                                    if (string.IsNullOrEmpty(BudgetWorkSheetSQL_FieldName))
                                                                        BudgetWorkSheetSQL_FieldName = k;
                                                                    else
                                                                        BudgetWorkSheetSQL_FieldName = BudgetWorkSheetSQL_FieldName + "," + k;
                                                                    */
                                                                    string[] location = new string[2];
                                                                    BudgetHeaderItemRowColDict.TryGetValue(k, out location);
                                                                    oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);


                                                                    if ((k.Contains("YrMth") || k.Contains("Date") || k.Contains("date")))
                                                                        if (oRng.get_Value(Type.Missing).GetType() == System.Type.GetType("System.DateTime"))
                                                                            command.Parameters.Add(new SqlParameter("@" + k, Convert.ToDateTime(oRng.get_Value(Type.Missing)).ToString("yyyy/MM/dd HH:mm:ss")));
                                                                        else
                                                                        {
                                                                            if (string.IsNullOrEmpty(oRng.get_Value(Type.Missing).ToString()))
                                                                                command.Parameters.Add(new SqlParameter("@" + k, System.DBNull.Value));
                                                                            else
                                                                                if (oRng.get_Value(Type.Missing).ToString() == "N/A")
                                                                                    command.Parameters.Add(new SqlParameter("@" + k, System.DBNull.Value));
                                                                                else
                                                                                    command.Parameters.Add(new SqlParameter("@" + k, oRng.get_Value(Type.Missing).ToString()));
                                                                        }
                                                                    else
                                                                        if (oRng.Value2 == null)
                                                                            command.Parameters.Add(new SqlParameter("@" + k, System.DBNull.Value));
                                                                        else
                                                                        {
                                                                            decimal decimal_out = 0;
                                                                            if (decimal.TryParse(oRng.Value2.ToString().Replace("'", "''"), out decimal_out) == true)
                                                                            {
                                                                                if (k.Contains("Budget_Project_No"))
                                                                                    Budget_Project_No = oRng.Value2.ToString().Replace("'", "''");
                                                                                command.Parameters.Add(new SqlParameter("@" + k, decimal_out));
                                                                            }
                                                                            else
                                                                                if (k.Contains("Hit_Rate"))
                                                                                    command.Parameters.Add(new SqlParameter("@" + k, decimal_out));
                                                                                else
                                                                                {
                                                                                    if (k.Contains("Budget_Project_No"))
                                                                                        Budget_Project_No = oRng.Value2.ToString().Replace("'", "''");
                                                                                    command.Parameters.Add(new SqlParameter("@" + k, oRng.Value2.ToString().Replace("'", "''")));
                                                                                }
                                                                        }

                                                                    //if (k == "CumTo_ForPeriod_YrMth") // debug
                                                                    //Console.WriteLine("");
                                                                    /*

                                                                    if (string.IsNullOrEmpty(BudgetWorkSheetSQL_Value))
                                                                        if ((k.Contains("YrMth") || k.Contains("Date") || k.Contains("date")))
                                                                            if (oRng.get_Value(Type.Missing).GetType() == System.Type.GetType("System.DateTime"))
                                                                                BudgetWorkSheetSQL_Value = "'" + Convert.ToDateTime(oRng.get_Value(Type.Missing)).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                                                                            else
                                                                                BudgetWorkSheetSQL_Value = "NULL";
                                                                        else
                                                                            if (oRng.Value2 == null)
                                                                                BudgetWorkSheetSQL_Value = "NULL";
                                                                            else
                                                                            {
                                                                                decimal decimal_out = 0;
                                                                                if (decimal.TryParse(oRng.Value2.ToString().Replace("'", "''"), out decimal_out) == true)
                                                                                    BudgetWorkSheetSQL_Value = oRng.Value2.ToString().Replace("'", "''");
                                                                                else
                                                                                    BudgetWorkSheetSQL_Value = "'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                            }
                                                                    else
                                                                        if ((k.Contains("YrMth") || k.Contains("Date") || k.Contains("date")))
                                                                            if (oRng.get_Value(Type.Missing).GetType() == System.Type.GetType("System.DateTime"))
                                                                                BudgetWorkSheetSQL_Value = BudgetWorkSheetSQL_Value + ",'" + Convert.ToDateTime(oRng.get_Value(Type.Missing)).ToString("yyyy/MM/dd HH:mm:ss") + "'";
                                                                            else
                                                                                BudgetWorkSheetSQL_Value = BudgetWorkSheetSQL_Value + ",NULL";
                                                                        else
                                                                            if (oRng.Value2 == null)
                                                                                BudgetWorkSheetSQL_Value = BudgetWorkSheetSQL_Value + ",NULL";
                                                                            else
                                                                            {
                                                                                decimal decimal_out = 0;
                                                                                if (decimal.TryParse(oRng.Value2.ToString().Replace("'", "''"), out decimal_out) == true)
                                                                                    BudgetWorkSheetSQL_Value = BudgetWorkSheetSQL_Value + "," + oRng.Value2.ToString().Replace("'", "''");
                                                                                else
                                                                                    BudgetWorkSheetSQL_Value = BudgetWorkSheetSQL_Value + ",'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                            }*/

                                                                }
                                                                //BudgetWorkSheetSQL = "insert into Budget_CashFlow_Worksheet (" + BudgetWorkSheetSQL_FieldName + ") values (" + BudgetWorkSheetSQL_Value + ")";
                                                                //ExecuteDatabase(BudgetWorkSheetSQL);

                                                                if (oSheet.Name == Budget_Project_No) // Check Sheet Name = Project No in Excel
                                                                {
                                                                    connection.Open();
                                                                    transa = connection.BeginTransaction();
                                                                    command.Transaction = transa;
                                                                    command.Connection = connection;
                                                                    command.ExecuteNonQuery();
                                                                    transa.Commit();

                                                                    Worksheet_ID = command.Parameters["@Budget_CashFlow_Worksheet_ID"].Value.ToString();
                                                                    string result = command.Parameters["@Result"].Value.ToString();

                                                                    command.Dispose();
                                                                    transa.Dispose();
                                                                    connection.Close();
                                                                }
                                                                once_worksheet = true;
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                sw.WriteLine(oSheet.Name + ": Stored Proc sp_Budget_CashFlow_Worksheet_Update error " + ex.Message.ToString());
                                                                number_of_error = number_of_error + 1;
                                                                connection.Close();
                                                            }
                                                        }
                                                    }

                                                    if (oSheet.Name == Budget_Project_No && number_of_error == 0) // Check Sheet Name = Project No in Excel
                                                    {
                                                        string Budget_CashFlow_Project_ID = "";
                                                        using (SqlConnection connection = new SqlConnection(_connectionString))
                                                        {
                                                            try
                                                            {
                                                                SqlTransaction transa = null;
                                                                SqlCommand command = new SqlCommand("sp_Budget_CashFlow_Project_Update", connection);
                                                                command.CommandType = CommandType.StoredProcedure;

                                                                int id = 0;

                                                                SqlParameter returnValue2 = new SqlParameter("@Budget_CashFlow_Project_ID", id); // ok
                                                                returnValue2.Direction = ParameterDirection.Output;

                                                                //SqlParameter returnValue2 = new SqlParameter("@Budget_CashFlow_Project_ID", id); // ok
                                                                //returnValue2.Direction = ParameterDirection.InputOutput;
                                                                command.Parameters.Add(returnValue2);
                                                                string r = "";
                                                                if (year_end_date.Month < 10)
                                                                    r = year_end_date.Year.ToString() + "0" + year_end_date.Month.ToString();
                                                                else
                                                                    r = year_end_date.Year.ToString() + year_end_date.Month.ToString();

                                                                //command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Worksheet_ID", "0"));
                                                                command.Parameters.Add(new SqlParameter("@YrEnd_CalPeriod", r));
                                                                command.Parameters.Add(new SqlParameter("@Budget_Project_No", Budget_Project_No));
                                                                command.Parameters.Add(new SqlParameter("@Last_Upload_Budget_CashFlow_Worksheet_ID", Worksheet_ID));

                                                                SqlParameter returnValue = new SqlParameter("@Result", ""); // ok
                                                                returnValue.Direction = ParameterDirection.Output;

                                                                command.Parameters.Add(returnValue);


                                                                connection.Open();
                                                                transa = connection.BeginTransaction();
                                                                command.Transaction = transa;
                                                                command.Connection = connection;
                                                                command.ExecuteNonQuery();
                                                                transa.Commit();

                                                                Budget_CashFlow_Project_ID = command.Parameters["@Budget_CashFlow_Project_ID"].Value.ToString();
                                                                string result = command.Parameters["@Result"].Value.ToString();

                                                                command.Dispose();
                                                                transa.Dispose();
                                                                connection.Close();
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                sw.WriteLine(oSheet.Name + ": Stored Proc sp_Budget_CashFlow_Project_Update error " + ex.Message.ToString());
                                                                number_of_error = number_of_error + 1;
                                                                connection.Close();
                                                            }
                                                        }
                                                        string BudgetDetailSQL = "Insert into ";
                                                        string BudgetDetailSQL_FieldName = "";
                                                        string BudgetDetailSQL_Value = "";
                                                        // Insert into

                                                        foreach (string[] k in BudgetDetailDict.Keys)
                                                        {
                                                            //id[0] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                                            //id[1] = dr["Budget_CashFlow_ReportItem_Name"].ToString();
                                                            //value[0] = dr["Row_No"].ToString();
                                                            //value[1] = dr["Budget_CashFlow_ReportItem_ID"].ToString();
                                                            //value[2] = dr["Budget_CashFlow_ReportItem_Code"].ToString();

                                                            string[] row = BudgetDetailDict[k];
                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                BudgetDetailSQL_FieldName = "Budget_CashFlow_Worksheet_ID";
                                                            else
                                                                BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + ",Budget_CashFlow_Worksheet_ID";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                BudgetDetailSQL_FieldName = "Budget_CashFlow_ReportItem_Name";
                                                            else
                                                                BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + ",Budget_CashFlow_ReportItem_Name";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                BudgetDetailSQL_FieldName = "Budget_CashFlow_ReportItem_ID";
                                                            else
                                                                BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + ",Budget_CashFlow_ReportItem_ID";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                BudgetDetailSQL_FieldName = "Budget_CashFlow_ReportItem_Code";
                                                            else
                                                                BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + ",Budget_CashFlow_ReportItem_Code";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                BudgetDetailSQL_FieldName = "Budget_CashFlow_Type";
                                                            else
                                                                BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + ",Budget_CashFlow_Type";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                                BudgetDetailSQL_Value = Worksheet_ID.Replace("'", "''");
                                                            else
                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + "," + Worksheet_ID.Replace("'", "''");

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                                BudgetDetailSQL_Value = "'" + k[1].Replace("'", "''") + "'";
                                                            else
                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + k[1].Replace("'", "''") + "'";

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                                BudgetDetailSQL_Value = "'" + row[1].Replace("'", "''") + "'";
                                                            else
                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + row[1].Replace("'", "''") + "'";

                                                            if (Worksheet_ID == "1")
                                                                Console.Write("");
                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                            {
                                                                if (row[2].Replace("'", "''") == "COMPULSORY" || row[2].Replace("'", "''") == "NO" || row[2].Replace("'", "''") == "WARNING")
                                                                {
                                                                    //buster2
                                                                    int count = 0;
                                                                    foreach (string k2 in Budget_CashFlow_ReportItem_CodeDict.Keys)
                                                                    {
                                                                        if (k2 == "Budget_CashFlow_ReportItem_Code")
                                                                        {
                                                                            string[] data = new string[2];
                                                                            Budget_CashFlow_ReportItem_CodeDict.TryGetValue(k2, out data);
                                                                            string col = data[0];

                                                                            if (!string.IsNullOrEmpty(col))
                                                                            {
                                                                                oRng = oSheet.get_Range(col + row[0], col + row[0]);
                                                                                if (Worksheet_ID == "1")
                                                                                    Console.Write("");
                                                                                if (oRng.Value2 != null)
                                                                                    BudgetDetailSQL_Value = "'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                                else
                                                                                    BudgetDetailSQL_Value = "NULL";
                                                                            }
                                                                            else
                                                                                BudgetDetailSQL_Value = "NULL";
                                                                            count = count + 1;
                                                                        }
                                                                    }
                                                                    if (count == 0)
                                                                        BudgetDetailSQL_Value = "NULL";
                                                                }
                                                                else
                                                                    BudgetDetailSQL_Value = "'" + row[2].Replace("'", "''") + "'";
                                                            }
                                                            else
                                                            {
                                                                if (row[2].Replace("'", "''") == "COMPULSORY" || row[2].Replace("'", "''") == "NO" || row[2].Replace("'", "''") == "WARNING")
                                                                {
                                                                    //buster2
                                                                    int count = 0;
                                                                    foreach (string k2 in Budget_CashFlow_ReportItem_CodeDict.Keys)
                                                                    {
                                                                        if (k2 == "Budget_CashFlow_ReportItem_Code")
                                                                        {
                                                                            string[] data = new string[2];
                                                                            Budget_CashFlow_ReportItem_CodeDict.TryGetValue(k2, out data);
                                                                            string col = data[0];

                                                                            if (!string.IsNullOrEmpty(col))
                                                                            {
                                                                                oRng = oSheet.get_Range(col + row[0], col + row[0]);
                                                                                if (Worksheet_ID == "1")
                                                                                    Console.Write("");
                                                                                if (oRng.Value2 != null)
                                                                                    BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                                else
                                                                                    BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",NULL";
                                                                            }
                                                                            else
                                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",NULL";
                                                                            count = count + 1;
                                                                        }
                                                                    }
                                                                    if (count == 0)
                                                                        BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",NULL";
                                                                }
                                                                else
                                                                    BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + row[2].Replace("'", "''") + "'";
                                                            }

                                                            if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                                BudgetDetailSQL_Value = "'" + row[3].Replace("'", "''") + "'";
                                                            else
                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + row[3].Replace("'", "''") + "'";

                                                            foreach (string k2 in BudgetDetailColDict.Keys)
                                                            {
                                                                string[] data = new string[2];
                                                                BudgetDetailColDict.TryGetValue(k2, out data);
                                                                string col = data[0];

                                                                if (k2 == "ForPeriod21") Console.WriteLine("");
                                                                if (string.IsNullOrEmpty(BudgetDetailSQL_FieldName))
                                                                    BudgetDetailSQL_FieldName = k2;
                                                                else
                                                                    BudgetDetailSQL_FieldName = BudgetDetailSQL_FieldName + "," + k2;

                                                                if (k2.StartsWith("CumTo_ForPeriod"))
                                                                    Console.WriteLine("");

                                                                if (!string.IsNullOrEmpty(col))
                                                                {
                                                                    oRng = oSheet.get_Range(col + row[0], col + row[0]);

                                                                    if (string.IsNullOrEmpty(BudgetDetailSQL_Value))
                                                                        if (oRng.Value2 == null)
                                                                            BudgetDetailSQL_Value = "NULL";
                                                                        else
                                                                        {
                                                                            if (oRng.Value2.ToString().StartsWith("Gross Value for Own Works"))
                                                                                Console.WriteLine("");
                                                                            if (oRng.Value2 != null)
                                                                                BudgetDetailSQL_Value = "'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                        }
                                                                    else
                                                                        if (oRng.Value2 == null)
                                                                            BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",NULL";
                                                                        else
                                                                        {
                                                                            if (oRng.Value2.ToString().StartsWith("Gross Value for Own Works"))
                                                                                Console.WriteLine("");
                                                                            if (oRng.Value2 != null)
                                                                                BudgetDetailSQL_Value = BudgetDetailSQL_Value + ",'" + oRng.Value2.ToString().Replace("'", "''") + "'";
                                                                        }
                                                                }
                                                            }
                                                            BudgetDetailSQL = "Insert into Budget_CashFlow_Detail (" + BudgetDetailSQL_FieldName + ") values (" + BudgetDetailSQL_Value + ")";
                                                            string error = "";
                                                            if (number_of_error == 0)
                                                                error = ExecuteDatabase(BudgetDetailSQL);

                                                            if (!string.IsNullOrEmpty(error))
                                                            {
                                                                sw.WriteLine(oSheet.Name + ": Insert into Budget_CashFlow_Detail error " + error);
                                                                number_of_error = number_of_error + 1;
                                                            }
                                                            BudgetDetailSQL = "";
                                                            BudgetDetailSQL_FieldName = "";
                                                            BudgetDetailSQL_Value = "";
                                                        }
                                                        //Console.WriteLine("");
                                                        using (SqlConnection connection = new SqlConnection(_connectionString))
                                                        {
                                                            try
                                                            {
                                                                SqlTransaction transa = null;
                                                                SqlCommand command = new SqlCommand("sp_Update_Worksheet_Status", connection);
                                                                command.CommandType = CommandType.StoredProcedure;

                                                                command.Parameters.Add(new SqlParameter("@Budget_CashFlow_Worksheet_ID", Worksheet_ID)); // ok
                                                                if (number_of_error == 0)
                                                                    command.Parameters.Add(new SqlParameter("@Upload_Status", "SUCCESS"));//ok
                                                                else
                                                                    command.Parameters.Add(new SqlParameter("@Upload_Status", "FAILED"));//ok

                                                                SqlParameter returnValue = new SqlParameter("@Result", ""); // ok
                                                                returnValue.Direction = ParameterDirection.Output;

                                                                command.Parameters.Add(returnValue);


                                                                connection.Open();
                                                                transa = connection.BeginTransaction();
                                                                command.Transaction = transa;
                                                                command.Connection = connection;
                                                                command.ExecuteNonQuery();
                                                                transa.Commit();

                                                                //Budget_CashFlow_Project_ID = command.Parameters["@Budget_CashFlow_Worksheet_ID"].Value.ToString();
                                                                string result = command.Parameters["@Result"].Value.ToString();

                                                                command.Dispose();
                                                                transa.Dispose();
                                                                connection.Close();
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                sw.WriteLine(oSheet.Name + ": Update database status error " + ex.Message.ToString());
                                                                connection.Close();
                                                            }
                                                        }
                                                        if (number_of_error > 0)
                                                            error_sheet.Add(error_sheet.Count.ToString(), oSheet.Name);
                                                    }
                                                }
                                            }
                                        }
                                        string errorMsg = "";
                                        foreach (string k2 in error_sheet.Keys)
                                        {
                                            string sheetname = "";
                                            error_sheet.TryGetValue(k2, out sheetname);
                                            if (string.IsNullOrEmpty(errorMsg))
                                                errorMsg = errorMsg + sheetname;
                                            else
                                                errorMsg = errorMsg + "," + sheetname;
                                        }
                                        if (error_sheet.Count == 0)
                                        {
                                            upload_status.Add(oSheet.Name, "SUCCESS");
                                            error_status.Add(oSheet.Name, "");

                                            if (isWrittenEndHtml == false)
                                            {
                                                //sw2.WriteLine("     <td>");
                                                //sw2.WriteLine("         SUCCESS");
                                                //sw2.WriteLine("     </td>");
                                                //sw2.WriteLine("     <td>");
                                                //sw2.WriteLine("         .");
                                                //sw2.WriteLine("     </td>");

                                            }
                                            //MessageBox.Show("Upload Success");
                                        }
                                        else
                                        {
                                            upload_status.Add(oSheet.Name, "FAILED");
                                            error_status.Add(oSheet.Name, "         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");

                                            if (isWrittenEndHtml == false)
                                            {
                                                //sw2.WriteLine("     <td>");
                                                //sw2.WriteLine("         FAILED");
                                                //sw2.WriteLine("     </td>");
                                                //sw2.WriteLine("     <td>");
                                                //sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                                //sw2.WriteLine("     </td>");

                                            }
                                            //MessageBox.Show("upload to database Error worksheet " + errorMsg);
                                        }

                                    }
                                    /*
                                    string errorMsg = "";
                                    foreach (string k2 in error_sheet.Keys)
                                    {
                                        string sheetname = "";
                                        error_sheet.TryGetValue(k2, out sheetname);
                                        if (string.IsNullOrEmpty(errorMsg))
                                            errorMsg = errorMsg + sheetname;
                                        else
                                            errorMsg = errorMsg + "," + sheetname;
                                    }
                                    if (error_sheet.Count == 0)
                                    {
                                        if (isWrittenEndHtml == false)
                                        {
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         SUCCESS");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            //sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            sw2.WriteLine("         .");
                                            sw2.WriteLine("     </td>");
                                        }
                                        //MessageBox.Show("Upload Success");
                                    }
                                    else
                                    {
                                        if (isWrittenEndHtml == false)
                                        {
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         FAILED");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            sw2.WriteLine("     </td>");
                                        }
                                        MessageBox.Show("upload to database Error worksheet " + errorMsg);
                                    }*/
                                }
                                else
                                {
                                    /*
                                    if (error_sheet.Count == 0)
                                    {
                                        if (isWrittenEndHtml == false)
                                        {
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         SUCCESS");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            //sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            sw2.WriteLine("         .");
                                            sw2.WriteLine("     </td>");
                                        }
                                        //MessageBox.Show("Upload Success");
                                    }
                                    else
                                    {
                                        if (isWrittenEndHtml == false)
                                        {
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         FAILED");
                                            sw2.WriteLine("     </td>");
                                            sw2.WriteLine("     <td>");
                                            sw2.WriteLine("         Please refer to <a href=\"" + error_path + "\">" + error_filename + "</a>");
                                            sw2.WriteLine("     </td>");
                                        }
                                        //MessageBox.Show("upload to database Error worksheet " + errorMsg);
                                    }*/
                                }
                            }
                            /*
                            if (oSheet.Cells[3, 1] == "Project Name")
                            {
                                oRng = oSheet.get_Range("A1", "A65536");
                                oRng.ColumnWidth = 114;
                            }

                            oRng = oSheet.get_Range("A1", "A65536");
                            //oXL.Visible = true;
                            */


                            oWB.Close(null, null, null);
                            oXL.Workbooks.Close();
                            oXL.Quit();
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(oRng);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB);
                            oSheet = null;
                            oWB = null;
                            oXL = null;
                            GC.Collect();

                            //Application.Exit();
                        }
                        catch (Exception ex)
                        {
                            //Error_Label.Text = "No file generated! " + " Error Occur - Generation of Excel " + ex.Message;
                            oWB.Close(null, null, null);
                            oXL.Workbooks.Close();
                            oXL.Quit();
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(oRng);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB);
                            oSheet = null;
                            oWB = null;
                            oXL = null;
                            GC.Collect();
                        }
                        #endregion many
                    }
                }
                Boolean one_of_failed = false;
                foreach (String k in log_sheet_name.Keys)
                {
                    String dummy = "";
                    if (upload_status.TryGetValue(k, out dummy) == false)
                    {
                        upload_status.Add(k, "FAILED");
                    }

                    dummy = "";
                    if (error_status.TryGetValue(k, out dummy) == false)
                    {
                        error_status.Add(k, "");
                    }
                }
                foreach (String k in log_sheet_name.Keys)
                {
                    String dummy = "";
                    sw2.WriteLine("<tr width=70>");
                    sw2.WriteLine("     <td>");
                    log_full_name.TryGetValue(k, out dummy);
                    sw2.WriteLine(dummy);
                    sw2.WriteLine("     </td>");
                    sw2.WriteLine("     <td>");
                    dummy = "";
                    log_sheet_name.TryGetValue(k, out dummy);
                    sw2.WriteLine(dummy);
                    sw2.WriteLine("     </td>");
                    sw2.WriteLine("     <td>");
                    dummy = "";
                    validation_status.TryGetValue(k, out dummy);
                    sw2.WriteLine(dummy);
                    sw2.WriteLine("     </td>");
                    sw2.WriteLine("     <td>");
                    dummy = "";
                    upload_status.TryGetValue(k, out dummy);
                    sw2.WriteLine(dummy);
                    sw2.WriteLine("     </td>");
                    sw2.WriteLine("     <td>");
                    dummy = "";
                    error_status.TryGetValue(k, out dummy);
                    sw2.WriteLine(dummy);
                    sw2.WriteLine("     </td>");
                    sw2.WriteLine("</tr>");
                }
                sw2.WriteLine("</table>");
                sw2.WriteLine("</html>");
                for(int i=0; i<10; i++)
                    progressBar1.PerformStep();
            }
            //Console.ReadKey();
            if (!string.IsNullOrEmpty(param[2]))
                if (File.Exists(param[2] + "\\Result.html"))
                    System.Diagnostics.Process.Start(param[2] + "\\Result.html");
            Application.Exit();
        }
        static public string ExecuteDatabase(string TempStoreStr)
        {
            SqlTransaction trans = null;
            string _connectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;

                    command.CommandText = TempStoreStr;

                    connection.Open();
                    trans = connection.BeginTransaction();
                    command.Transaction = trans;
                    command.Connection = connection;
                    command.ExecuteNonQuery();
                    trans.Commit();

                    trans.Dispose();
                    connection.Close();
                    command.Dispose();
                    connection.Dispose();
                }
                catch (Exception ex)
                {
                    trans.Rollback();
                    trans.Dispose();
                    connection.Close();
                    connection.Dispose();
                    return ex.Message;
                }
            }
            return "";
        }
        Dictionary<string, string[]> ReadBudget_CashFlow_ReportItem_CodeFromDB(string cashflow_type)
        {
            Dictionary<string, string[]> result = new Dictionary<string, string[]>();
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    //command.CommandText = "select DataField_Name, Col_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','CumTo_ForPeriod','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "'";
                    command.CommandText = "select DataField_Name, Col_No, DataField_Type from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name = 'Budget_CashFlow_ReportItem_Code' and Budget_CashFlow_Type='" + cashflow_type + "'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["DataField_Name"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                            {
                                string[] data = new string[2];
                                data[0] = dr["Col_No"].ToString();
                                if (dr["DataField_Type"] != System.DBNull.Value)
                                    data[1] = dr["DataField_Type"].ToString();
                                else
                                    data[1] = "";
                                result.Add(dr["DataField_Name"].ToString(), data);
                            }
                        }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }
        Dictionary<string, string[]> ReadBudget_CashFlow_AllReportItem_CodeFromDB(string cashflow_type)
        {
            Dictionary<string, string[]> result = new Dictionary<string, string[]>();
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    //command.CommandText = "select DataField_Name, Col_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','CumTo_ForPeriod','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "'";
                    command.CommandText = "select DataField_Name, Col_No, DataField_Type from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name', 'Budget_CashFlow_ReportItem_Code') and Budget_CashFlow_Type='" + cashflow_type + "'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["DataField_Name"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                            {
                                string[] data = new string[2];
                                data[0] = dr["Col_No"].ToString();
                                if (dr["DataField_Type"] != System.DBNull.Value)
                                    data[1] = dr["DataField_Type"].ToString();
                                else
                                    data[1] = "";
                                result.Add(dr["DataField_Name"].ToString(), data);
                            }
                        }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }
        /*
        static string StartRow(string Budget_CashFlow_Type, string latestperiod)
        {
            string result = "";
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    command.CommandText = "select YrEnd_CalPeriod,Budget_CashFlow_Type,Row_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name='[START_ROW_NO]' and Budget_CashFlow_Type='" + Budget_CashFlow_Type.ToUpper() + "' and YrEnd_CalPeriod='" + latestperiod + "'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Row_No"] != System.DBNull.Value)
                        {
                            result = dr["Row_No"].ToString();
                        }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }*/
        static string ProjectCodeExistInPCMS(string code)
        {
            string result = "";
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PCMSConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    command.CommandText = "select count(*) as num from OPRJ where PrjCode='"+code+"'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["num"] != System.DBNull.Value)
                        {
                            result = dr["num"].ToString();
                        }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }
        static string CostCodeExistInPCMS(string code)
        {
            string result = "";
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PCMSConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    command.CommandText = "select count(*) as num from [@RPTCODE] where Code='" + code + "'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["num"] != System.DBNull.Value)
                        {
                            result = dr["num"].ToString();
                        }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }
        static string CashFlowCodeExistInPCMS(string code)
        {
            string result = "";
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PCMSConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    command.CommandText = "select count(*) as num from [@CASHCODE] where Code='" + code + "'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["num"] != System.DBNull.Value)
                        {
                            result = dr["num"].ToString();
                        }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }
        static string LatestPeriodInHeader()
        {
            string result = "";
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    command.CommandText = "select distinct YrEnd_CalPeriod from dbo.Budget_CashFlow_Excel_HeaderItem_Detail order by YrEnd_CalPeriod desc";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["YrEnd_CalPeriod"] != System.DBNull.Value)
                        {
                            result = dr["YrEnd_CalPeriod"].ToString();
                        }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }
        static string LatestPeriodInReport()
        {
            string result = "";
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    command.CommandText = "select distinct YrEnd_CalPeriod from dbo.Budget_CashFlow_Excel_ReportItem_Detail order by YrEnd_CalPeriod desc";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["YrEnd_CalPeriod"] != System.DBNull.Value)
                        {
                            result = dr["YrEnd_CalPeriod"].ToString();
                        }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }
        static int[] GetLocation(Excel._Worksheet oSheet, string target, Boolean secondSearch)
        {
            int[] sheetlocation = new int[2];
            sheetlocation[0] = 0;//row
            sheetlocation[1] = 0;//col

            Excel.Range oRng = oSheet.get_Range("A1", "AZ65536");

            Excel.Range currentFind = null;
            object missing = System.Type.Missing;

            currentFind = oRng.Find(target, missing,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, missing, missing);

            if (currentFind != null)
            {
                sheetlocation[0] = currentFind.Row;
                sheetlocation[1] = currentFind.Column;

                if (secondSearch == true)
                {
                    currentFind = oRng.Find(target, missing,
                    Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, missing, missing);

                    if (currentFind != null)
                    {
                        sheetlocation[0] = currentFind.Row;
                        sheetlocation[1] = currentFind.Column;
                    }
                }
            }
            return sheetlocation;
        }
        static Dictionary<string, string[]> ReadHeaderFromDB(string cashflow_type)
        {
            Dictionary<string, string[]> result = new Dictionary<string, string[]>();
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    command.CommandText = "select Upload_Excel_Field_Name, Col_No, Row_No from Budget_CashFlow_Excel_HeaderItem_Detail where Budget_CashFlow_Type = '" + cashflow_type + "'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Upload_Excel_Field_Name"] != System.DBNull.Value)
                            if (dr["Col_No"] != System.DBNull.Value)
                                if (dr["Row_No"] != System.DBNull.Value)
                                {
                                    string[] location = new string[2];
                                    location[0] = dr["Row_No"].ToString();
                                    location[1] = dr["Col_No"].ToString();
                                    result.Add(dr["Upload_Excel_Field_Name"].ToString(), location);
                                }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }
        static Dictionary<string, string[]> ReadDetailColFromDB(string cashflow_type)
        {
            Dictionary<string, string[]> result = new Dictionary<string, string[]>();
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    //command.CommandText = "select DataField_Name, Col_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','CumTo_ForPeriod','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "'";
                    command.CommandText = "select DataField_Name, Col_No, DataField_Type from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "' and Heading_Type = 'D'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["DataField_Name"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                            {
                                string[] data = new string[2];
                                data[0] = dr["Col_No"].ToString();
                                if (dr["DataField_Type"] != System.DBNull.Value)
                                    data[1] = dr["DataField_Type"].ToString();
                                else
                                    data[1] = "";
                                result.Add(dr["DataField_Name"].ToString(), data);
                            }
                        }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }
        static Dictionary<string[], string[]> ReadDetailFromDB(string cashflow_type, string Budget_Project_On_Hand_Status)
        {
            Dictionary<string[], string[]> result = new Dictionary<string[], string[]>();
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    // buster
                    command.CommandText = "select Budget_CashFlow_Type, Budget_CashFlow_ReportItem_Name, Check_Validation_Code_Input, Row_No, Budget_CashFlow_ReportItem_ID, Budget_CashFlow_Excel_ReportItem_Detail_ID, Budget_CashFlow_ReportItem_Code from dbo.Budget_CashFlow_Excel_ReportItem_Detail where Heading_Type = 'D' and Budget_CashFlow_Type='" + cashflow_type + "' and Budget_Project_On_Hand_Status = '" + Budget_Project_On_Hand_Status + "'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Budget_CashFlow_ReportItem_Name"] != System.DBNull.Value)
                        {
                            if (dr["Row_No"] != System.DBNull.Value)
                            {
                                string[] id = new string[2];
                                id[0] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                id[1] = dr["Budget_CashFlow_ReportItem_Name"].ToString();

                                string[] value = new string[4];
                                value[0] = dr["Row_No"].ToString();
                                value[1] = dr["Budget_CashFlow_ReportItem_ID"].ToString();
                                if (string.IsNullOrEmpty(dr["Budget_CashFlow_ReportItem_Code"].ToString()))
                                    value[2] = dr["Check_Validation_Code_Input"].ToString();
                                else
                                    value[2] = dr["Budget_CashFlow_ReportItem_Code"].ToString();
                                value[3] = dr["Budget_CashFlow_Type"].ToString();
                                result.Add(id, value);
                            }
                        }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }

        static string[] ReadBudget_Project_On_Hand_StatusFromDB(string latestperiod)
        {
            string[] location = new string[2];
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    command.CommandText = "select Budget_Project_On_Hand_Status,Col_No,Row_No from dbo.Budget_CashFlow_Excel_HeaderItem_RowCol where YrEnd_CalPeriod='" + latestperiod + "' and DataField_Name='Budget_Project_On_Hand_Status'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Col_No"] != System.DBNull.Value)
                            if (dr["Row_No"] != System.DBNull.Value)
                            {
                                location[0] = dr["Row_No"].ToString();
                                location[1] = dr["Col_No"].ToString();
                            }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return location;
        }

        static Dictionary<string, string[]> ReadHeaderItemRowColFromDB(string Budget_Project_On_Hand_Status, string latestperiod)
        {
            Dictionary<string, string[]> result = new Dictionary<string, string[]>();
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    command.CommandText = "select DataField_Name, Col_No, Row_No, DataField_Type from dbo.Budget_CashFlow_Excel_HeaderItem_RowCol where Budget_Project_On_Hand_Status = '" + Budget_Project_On_Hand_Status + "' and YrEnd_CalPeriod='" + latestperiod + "'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["DataField_Name"] != System.DBNull.Value)
                            if (dr["Col_No"] != System.DBNull.Value)
                                if (dr["Row_No"] != System.DBNull.Value)
                                {
                                    string[] location = new string[3];
                                    location[0] = dr["Row_No"].ToString();
                                    location[1] = dr["Col_No"].ToString();
                                    if (dr["DataField_Type"] != System.DBNull.Value)
                                        location[2] = dr["DataField_Type"].ToString();    
                                    result.Add(dr["DataField_Name"].ToString(), location);
                                }
                    }
                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            }
            return result;
        }
    }
}
