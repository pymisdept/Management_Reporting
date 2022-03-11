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

namespace OutputCashFlowExcel
{
    public partial class Form1 : Form
    {
        public Form1(string[] param)
        {
            InitializeComponent();

            label3.Visible = false;

            if (!string.IsNullOrEmpty(param[2]) && !string.IsNullOrEmpty(param[3]))
            {
                label3.Visible = true;
                try
                {
                    //this.Show();
                    //MessageBox.Show("before thread");
                    Thread newThread = new Thread(DoWork);

                    object args = new object[4] { param[0], param[1], param[2], param[3] };
                    //MessageBox.Show("before start thread");
                
                    newThread.Start(args);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }

            /*
            System.Drawing.Font currentFont = richTextBox1.SelectionFont;
            System.Drawing.FontStyle newFontStyle;
            newFontStyle = FontStyle.Regular;

            richTextBox1.Text = "A";
            richTextBox1.Font = new Font(currentFont.FontFamily, 12, newFontStyle);
            richTextBox1.Text = richTextBox1.Text + "2";
            richTextBox1.SelectionStart = 1;
            richTextBox1.SelectionLength = 1;
            richTextBox1.SelectionFont = new Font(currentFont.FontFamily,8,newFontStyle);
            richTextBox1.SelectionCharOffset = 5;
            //richTextBox1.SelectionCharOffset = -5;
            */

        }
        public void DoWork(object args)
        {
            //MessageBox.Show("before do work");
            Array argArray = new object[4];
            argArray = (Array)args;
            string[] param = new string[4];
            param[0] = (string)argArray.GetValue(0);
            param[1] = (string)argArray.GetValue(1);
            param[2] = (string)argArray.GetValue(2);
            param[3] = (string)argArray.GetValue(3);
            //MessageBox.Show("before do work 2");

            if (!string.IsNullOrEmpty(param[2]) && !string.IsNullOrEmpty(param[3]))
            {
                Export_button.Visible = false;
                ProjectCodeFrom_comboBox.Visible = false;
                ProjectCodeTo_comboBox.Visible = false;
                Date_comboBox.Visible = false;
                label1.Visible = false;
                label2.Visible = false;

                List<String> Project_Code_List = new List<String>();
                string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
                using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
                {
                    try
                    {
                        SqlCommand command = new SqlCommand();
                        command.CommandType = CommandType.Text;
                        //command.CommandText = "select DataField_Name, Col_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','CumTo_ForPeriod','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "'";
                        command.CommandText = "select Budget_Project_No from dbo.Budget_CashFlow_Project where Budget_Project_No >= '" + param[2] + "' and Budget_Project_No <= '" + param[3] + "' and YrEnd_CalPeriod = '" + param[1] + "'";
                        connection.Open();
                        command.Connection = connection;
                        SqlDataReader dr = command.ExecuteReader();
                        while (dr.Read())
                        {
                            if (dr["Budget_Project_No"] != System.DBNull.Value)
                            {
                                Project_Code_List.Add(dr["Budget_Project_No"].ToString());
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

                object Missing = System.Type.Missing;
                Excel._Workbook oWB = null;
                Excel._Worksheet oSheet = null;
                Excel.Range oRng = null;
                Excel.Application oXL = new Excel.Application();
                Excel.Sheets xlSheets = null;
                try
                {
                    oXL.Visible = false;
                    oXL.DisplayAlerts = false;

                    //Get a new workbook.
                    oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing));
                    Boolean first = true;

                    foreach (String project_code_item in Project_Code_List)
                    {
                        if (first == true)
                        {
                            oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                            first = false;
                        }
                        else
                        {
                            xlSheets = oWB.Sheets as Excel.Sheets;
                            oSheet = (Excel._Worksheet)xlSheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                            oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                        }
                        oSheet.Name = project_code_item;
                        label3.Text = project_code_item;
                        Dictionary<string, Dictionary<string, string[]>> Detail = ReadDetailRowFromDB(project_code_item, param[1]);

                        //Dictionary<string, Dictionary<string, string[]>> DetailIncludingColTotal = ReadDetailIncludingColTotalRowFromDB(ProjectCode_comboBox.SelectedValue.ToString());

                        Dictionary<string, string> DetailCol = null;
                        Dictionary<string, string[]> DetailCol2 = null;
                        bool once = false;
                        // Display the ProgressBar control.
                        progressBar1.Visible = true;
                        // Set Minimum to 1 to represent the first file being copied.
                        progressBar1.Minimum = 1;
                        // Set Maximum to the total number of files to copy.
                        progressBar1.Maximum = Detail.Count + 10;
                        // Set the initial value of the ProgressBar.
                        progressBar1.Value = 1;
                        // Set the Step property to a value of 1 to represent each file being copied.
                        progressBar1.Step = 1;

                        string YearEnd = "";
                        foreach (string Budget_CashFlow_ReportItem_Name in Detail.Keys)
                        {
                            Dictionary<string, string[]> data;
                            Detail.TryGetValue(Budget_CashFlow_ReportItem_Name, out data);

                            foreach (string column_name in data.Keys)
                            {
                                string[] location_and_value;
                                data.TryGetValue(column_name, out location_and_value);
                                string row = location_and_value[0];
                                string value = location_and_value[1];
                                string type = location_and_value[2];
                                string yrend_period = location_and_value[3];

                                if (once == false)
                                {
                                    YearEnd = yrend_period;
                                    DetailCol = ReadDetailColFromDB(type, yrend_period);
                                    once = true;
                                }
                                /*
                                if (type == "H")
                                {
                                    oRng = oSheet.get_Range(header_col + row, header_col + row);
                                    oRng.set_Value(Missing, column_name);
                                }
                                */
                                string col = "";
                                if (DetailCol.Count > 0)
                                    DetailCol.TryGetValue(column_name, out col);
                                if (col == "I")
                                    Console.WriteLine("");
                                if (column_name == "ForPeriod21")
                                    Console.WriteLine("");
                                if (!string.IsNullOrEmpty(col) && !string.IsNullOrEmpty(row))
                                {
                                    oSheet.Cells[row, 1] = Budget_CashFlow_ReportItem_Name.Substring(0, Budget_CashFlow_ReportItem_Name.IndexOf("$"));

                                    if (string.IsNullOrEmpty(value))
                                        oSheet.Cells[row, col] = System.String.Format("0", "##,###,###,##0.00");
                                    else
                                        oSheet.Cells[row, col] = System.String.Format(value, "##,###,###,##0.00");
                                    oRng = (Excel.Range)oSheet.Cells[row, col];
                                    oRng.NumberFormatLocal = "#,##0.00_);(#,##0.00)";
                                    oRng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                }
                            }
                            progressBar1.PerformStep();
                        }
                        // Column Total added at 27 Jan 2012
                        once = false;
                        //Dictionary<string, string[]> DetailCol2 = null;
                        foreach (string Budget_CashFlow_ReportItem_Name in Detail.Keys)
                        {
                            Dictionary<string, string[]> data;
                            Detail.TryGetValue(Budget_CashFlow_ReportItem_Name, out data);

                            foreach (string column_name in data.Keys)
                            {
                                string[] location_and_value;
                                data.TryGetValue(column_name, out location_and_value);
                                string row = location_and_value[0];
                                string value = location_and_value[1];
                                string type = location_and_value[2];
                                string yrend_period = location_and_value[3];

                                if (once == false)
                                {
                                    YearEnd = yrend_period;
                                    DetailCol2 = ReadDetailColTotalFromDB(type, yrend_period);
                                    once = true;
                                }
                                /*
                                if (type == "H")
                                {
                                    oRng = oSheet.get_Range(header_col + row, header_col + row);
                                    oRng.set_Value(Missing, column_name);
                                }
                                */

                                //if( DetailCol.Count > 0)
                                //DetailCol.TryGetValue(column_name, out col);
                                string col = "";
                                string typ = "";
                                string formula = "";
                                foreach (string k in DetailCol2.Keys)
                                {
                                    string[] col_and_formula = new string[2];
                                    DetailCol2.TryGetValue(k, out col_and_formula);
                                    col = col_and_formula[0];
                                    formula = col_and_formula[1];
                                    typ = col_and_formula[2];

                                    if (typ == "T")
                                    {
                                        string[] temp = formula.Replace("[", "").Replace("]", "").Split('+');

                                        string col2 = "";
                                        string typ2 = "";
                                        string formula2 = "";
                                        string final_formula = "=SUM(";
                                        foreach (string k2 in DetailCol2.Keys)
                                        {
                                            string[] col_and_formula2 = new string[2];
                                            DetailCol2.TryGetValue(k2, out col_and_formula2);
                                            col2 = col_and_formula2[0];
                                            formula2 = col_and_formula2[1];
                                            typ2 = col_and_formula2[2];

                                            if (typ2 == "D" || typ2 == "T")
                                            {
                                                foreach (string id in temp)
                                                {
                                                    if (k2 == id)
                                                    {
                                                        if (final_formula.EndsWith("("))
                                                            final_formula = final_formula + col2 + row;
                                                        else
                                                            final_formula = final_formula + "," + col2 + row;
                                                    }
                                                }
                                            }
                                        }
                                        final_formula = final_formula + ")";
                                        if (!string.IsNullOrEmpty(col) && !string.IsNullOrEmpty(row))
                                        {
                                            //oSheet.Cells[row, 1] = Budget_CashFlow_ReportItem_Name.Substring(0, Budget_CashFlow_ReportItem_Name.IndexOf("$"));
                                            oRng = oSheet.get_Range(col + row, col + row);
                                            oRng.Formula = final_formula;
                                        }
                                    }
                                }
                            }
                        }

                        Dictionary<string, string> worksheet = ReadWorkSheetFromDB(project_code_item);
                        string Budget_Project_On_Hand_Status = "";
                        worksheet.TryGetValue("Budget_Project_On_Hand_Status", out Budget_Project_On_Hand_Status);
                        string YrEnd_Date = "";
                        worksheet.TryGetValue("YrEnd_Date", out YrEnd_Date);
                        int year = Convert.ToDateTime(YrEnd_Date).Year;
                        int month = Convert.ToDateTime(YrEnd_Date).Month;
                        int day = Convert.ToDateTime(YrEnd_Date).Day;
                        string Year_End = "";
                        if (month < 10)
                            Year_End = year.ToString() + "0" + month.ToString();
                        else
                            Year_End = year.ToString() + month.ToString();


                        if (!string.IsNullOrEmpty(Budget_Project_On_Hand_Status) && !string.IsNullOrEmpty(YrEnd_Date))
                        {
                            //Header
                            Dictionary<string, string[]> HeaderRowCol = ReadHeaderRowColFromDB(Year_End, Budget_Project_On_Hand_Status);

                            foreach (string HeaderFieldName in worksheet.Keys)
                            {
                                string[] location = new string[3];
                                HeaderRowCol.TryGetValue(HeaderFieldName, out location);
                                if (location != null)
                                {
                                    if (!string.IsNullOrEmpty(location[0]) && !string.IsNullOrEmpty(location[1]))
                                    {
                                        //if (location[1] + location[0] == "E10")
                                            //Console.WriteLine("");
                                        oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);
                                        string value = "";
                                        worksheet.TryGetValue(HeaderFieldName, out value);
                                        DateTime date_result;
                                        if (DateTime.TryParse(value, out date_result) && location[2] != null)
                                        {
                                            if (HeaderFieldName == "CumTo_ForPeriod_YrMth" ||
                                                HeaderFieldName == "ForPeriodC9_YrMth" ||
                                                HeaderFieldName == "ForPeriodC10_YrMth" ||
                                                HeaderFieldName == "ForPeriodC11_YrMth" ||
                                                HeaderFieldName == "ForPeriodC12_YrMth" ||
                                                HeaderFieldName == "ForPeriod1_YrMth" ||
                                                HeaderFieldName == "ForPeriod2_YrMth" ||
                                                HeaderFieldName == "ForPeriod3_YrMth" ||
                                                HeaderFieldName == "ForPeriod4_YrMth" ||
                                                HeaderFieldName == "ForPeriod5_YrMth" ||
                                                HeaderFieldName == "ForPeriod6_YrMth" ||
                                                HeaderFieldName == "ForPeriod7_YrMth" ||
                                                HeaderFieldName == "ForPeriod8_YrMth" ||
                                                HeaderFieldName == "ForPeriod9_YrMth" ||
                                                HeaderFieldName == "ForPeriod10_YrMth" ||
                                                HeaderFieldName == "ForPeriod11_YrMth" ||
                                                HeaderFieldName == "ForPeriod12_YrMth" ||
                                                HeaderFieldName == "ForPeriod13_YrMth" ||
                                                HeaderFieldName == "ForPeriod14_YrMth" ||
                                                HeaderFieldName == "ForPeriod15_YrMth" ||
                                                HeaderFieldName == "ForPeriod16_YrMth" ||
                                                HeaderFieldName == "ForPeriod17_YrMth" ||
                                                HeaderFieldName == "ForPeriod18_YrMth" ||
                                                HeaderFieldName == "ForPeriod19_YrMth" ||
                                                HeaderFieldName == "ForPeriod20_YrMth" ||
                                                HeaderFieldName == "ForPeriod21_YrMth" ||
                                                HeaderFieldName == "ForPeriod22_YrMth" ||
                                                HeaderFieldName == "ForPeriod23_YrMth" ||
                                                HeaderFieldName == "ForPeriod24_YrMth" ||
                                                HeaderFieldName == "ForPeriod_After_YrMth" ||
                                                HeaderFieldName == "ForPeriod_After_YrMth" ||
                                                HeaderFieldName == "ForPeriod_After_YrMth"
                                                )
                                            {
                                                string y = date_result.ToString("yy");
                                                string m = date_result.ToString("MMM");

                                                oRng.set_Value(Missing, m + "-" + y);
                                            }
                                            else
                                            {
                                                if (location[2] == "D")
                                                    oRng.set_Value(Missing, date_result.ToString("dd-MMM-yyyy"));
                                                else
                                                {
                                                    oRng.set_Value(Missing, value.ToString());
                                                    oRng.NumberFormatLocal = "#,##0.00_);(#,##0.00)";
                                                    oRng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            oRng.set_Value(Missing, value);
                                            if (Convert.ToInt32(location[0]) < 10)
                                            {
                                                oRng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                                oRng.NumberFormatLocal = "#,##0.00_);(#,##0.00)";
                                            }
                                        }
                                    }
                                }

                            }
                            // Header formula
                            Dictionary<string, string[]> HeaderRowColFormula = ReadHeaderRowColFormulaFromDB(Year_End, Budget_Project_On_Hand_Status);
                            int[] PeriodRow = ReadHeaderRowColPeriodRowFromDB(Year_End, Budget_Project_On_Hand_Status);

                            foreach (string id in HeaderRowColFormula.Keys)
                            {
                                string[] location = new string[3];
                                HeaderRowColFormula.TryGetValue(id, out location);
                                if (location != null)
                                {
                                    if (!string.IsNullOrEmpty(location[0]) && !string.IsNullOrEmpty(location[1]))
                                    {
                                        oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);

                                        Boolean found = false;
                                        for (int i = 0; i < PeriodRow.Count(); i++)
                                        {
                                            if (Convert.ToInt32(location[0]) == PeriodRow[i])
                                            {
                                                found = true;
                                            }
                                        }
                                        if (found == false)
                                        {
                                            oRng.set_Value(Missing, System.String.Format("0", "##,###,###,##0.00"));
                                            oRng.NumberFormatLocal = "#,##0.00_);(#,##0.00)";
                                        }

                                        //oRng.set_Value(Missing, System.String.Format("0", "##,###,###,##0.00"));
                                        //oRng.NumberFormatLocal = "#,##0.00_);(#,##0.00)";
                                        string value = "";
                                        oRng.Formula = location[2];
                                    }
                                }
                            }
                            /*
                            foreach (string HeaderFieldName in worksheet.Keys)
                            {
                                string[] location = new string[3];
                                HeaderRowCol.TryGetValue(HeaderFieldName, out location);
                                if (location != null)
                                {
                                    if (!string.IsNullOrEmpty(location[0]) && !string.IsNullOrEmpty(location[1]))
                                    {
                                        oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);
                                        string value = "";
                                        worksheet.TryGetValue(HeaderFieldName, out value);
                                        DateTime date_result;
                                        if (DateTime.TryParse(value, out date_result) && location[2] != null)
                                        {
                                            if (HeaderFieldName == "ForPeriodC9_YrMth"
                                                )
                                            {
                                                string y = date_result.ToString("yy");
                                                string m = date_result.ToString("MMM");

                                                oRng.set_Value(Missing, m + "-" + y);
                                            }
                                            else
                                                oRng.set_Value(Missing, date_result.ToString("dd-MMM-yyyy"));
                                        }
                                        else
                                            oRng.set_Value(Missing, value);
                                    }
                                }
                            }*/
                        }
                        Dictionary<string, string[]> ReportItem_Detail_Header = ReadReportItem_Detail_Header(Budget_Project_On_Hand_Status);
                        foreach (string HeaderFieldName in ReportItem_Detail_Header.Keys)
                        {
                            string[] headername_location = new string[2];
                            ReportItem_Detail_Header.TryGetValue(HeaderFieldName, out headername_location);
                            if (headername_location != null)
                            {
                                if (!string.IsNullOrEmpty(headername_location[0]) && !string.IsNullOrEmpty(headername_location[1]))
                                {
                                    oRng = oSheet.get_Range(headername_location[0] + headername_location[1], headername_location[0] + headername_location[1]);
                                    string value = "";
                                    oRng.set_Value(Missing, HeaderFieldName);
                                }
                            }
                        }
                        // Only deal with all cost except Total direct cost and Total indirect cost
                        Dictionary<string, string[]> ReportItem_Detail_Total = ReadReportItem_Detail_Total(Budget_Project_On_Hand_Status);
                        Dictionary<String, String> ForTotalConstructionCostCol = new Dictionary<String, String>();

                        List<Formula_Class> formulas = new List<Formula_Class>();
                        foreach (string HeaderFieldName in ReportItem_Detail_Total.Keys)
                        {
                            string[] headername_location = new string[4];
                            ReportItem_Detail_Total.TryGetValue(HeaderFieldName, out headername_location);
                            if (headername_location != null)
                            {
                                if (!string.IsNullOrEmpty(headername_location[0]) && !string.IsNullOrEmpty(headername_location[1]))
                                {
                                    oRng = oSheet.get_Range(headername_location[0] + headername_location[1], headername_location[0] + headername_location[1]);
                                    string value = "";
                                    oRng.set_Value(Missing, HeaderFieldName);
                                }
                                if (headername_location.Count() > 2)
                                {
                                    if (!string.IsNullOrEmpty(headername_location[2]))
                                    {
                                        string[] total_cells = headername_location[2].Split('+');
                                        for (int i = 0; i < total_cells.Count(); i++)
                                        {
                                            total_cells[i] = total_cells[i].Replace("[", "").Replace("]", "");
                                        }
                                        string formula = "";
                                        once = false;
                                        foreach (string Budget_CashFlow_ReportItem_Name in Detail.Keys)
                                        {
                                            Dictionary<string, string[]> data;
                                            Detail.TryGetValue(Budget_CashFlow_ReportItem_Name, out data);

                                            foreach (string column_name in data.Keys)
                                            {
                                                string[] location_and_value;
                                                data.TryGetValue(column_name, out location_and_value);
                                                string row = location_and_value[0];
                                                string value = location_and_value[1];
                                                string type = location_and_value[2];
                                                string yrend_period = location_and_value[3];
                                                string seq = location_and_value[4];

                                                if (column_name == "ForPeriod21")
                                                    Console.WriteLine("");
                                                if (headername_location[3] == type)
                                                {
                                                    if (HeaderFieldName == "Total Cash Inflow")
                                                        Console.WriteLine("");
                                                    if (once == false)
                                                    {
                                                        YearEnd = yrend_period;
                                                        DetailCol = ReadDetailColFromDB(type, yrend_period);
                                                        once = true;
                                                    }
                                                    string col = "";
                                                    if (DetailCol.Count > 0)
                                                        DetailCol.TryGetValue(column_name, out col);

                                                    if (col == "F")
                                                        Console.WriteLine("");
                                                    if (!string.IsNullOrEmpty(col))
                                                    {
                                                        if (total_cells[0] == seq)
                                                        {
                                                            Formula_Class fc = new Formula_Class();

                                                            formula = "=Sum(" + col + row;
                                                            fc.formula = formula;
                                                            fc.col = col;
                                                            fc.row = headername_location[1];
                                                            fc.total_name = HeaderFieldName;


                                                            Boolean found = false;
                                                            foreach (Formula_Class f in formulas)
                                                            {
                                                                double double_result = 0;
                                                                if (Double.TryParse(f.formula.Substring(6, 1), out double_result) == true)
                                                                {
                                                                    if (f.formula.Substring(5, 1) == col && f.total_name == HeaderFieldName)
                                                                        found = true;
                                                                }
                                                                else
                                                                {
                                                                    if (f.formula.Substring(5, 2) == col && f.total_name == HeaderFieldName)
                                                                        found = true;
                                                                }
                                                            }
                                                            if (found == false)
                                                            {
                                                                formulas.Add(fc);
                                                                if (!ForTotalConstructionCostCol.ContainsValue(col))
                                                                    ForTotalConstructionCostCol.Add(ForTotalConstructionCostCol.Count.ToString(), col);
                                                            }
                                                        }
                                                        if (total_cells[total_cells.Count() - 1] == seq)
                                                        {
                                                            foreach (Formula_Class fc in formulas)
                                                            {
                                                                double double_result = 0;
                                                                if (Double.TryParse(fc.formula.Substring(6, 1), out double_result) == true)
                                                                {
                                                                    if (fc.formula.Substring(5, 1) == col && fc.total_name == HeaderFieldName)
                                                                        if (!fc.formula.EndsWith(")"))
                                                                            fc.formula = fc.formula + ":" + col + row + ")";
                                                                }
                                                                else
                                                                {
                                                                    if (fc.formula.Substring(5, 2) == col && fc.total_name == HeaderFieldName)
                                                                        if (!fc.formula.EndsWith(")"))
                                                                            fc.formula = fc.formula + ":" + col + row + ")";
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        // Total's Total except total direct cost and indirect cost - extreme buster

                        foreach (string HeaderFieldName in ReportItem_Detail_Total.Keys)
                        {
                            string[] headername_location = new string[4];
                            ReportItem_Detail_Total.TryGetValue(HeaderFieldName, out headername_location);
                            if (headername_location != null)
                            {
                                if (!string.IsNullOrEmpty(headername_location[0]) && !string.IsNullOrEmpty(headername_location[1]))
                                {
                                    oRng = oSheet.get_Range(headername_location[0] + headername_location[1], headername_location[0] + headername_location[1]);
                                    string value = "";
                                    oRng.set_Value(Missing, HeaderFieldName);
                                }
                                if (headername_location.Count() > 2)
                                {
                                    if (!string.IsNullOrEmpty(headername_location[2]))
                                    {
                                        string[] total_cells = headername_location[2].Replace("[", "").Replace("]", "").Split('+');

                                        string formula = "";
                                        once = false;
                                        DetailCol2 = ReadDetailColTotalTotalFromDB("BUDGET", Year_End);

                                        Dictionary<string, string> rowinfo = ReadTotalTotalRowInfoFromDB(Year_End, Budget_Project_On_Hand_Status);
                                        foreach (string k in DetailCol2.Keys)
                                        {
                                            string[] col = new string[3];
                                            DetailCol2.TryGetValue(k, out col);

                                            Formula_Class fc = new Formula_Class();

                                            formula = "=Sum(";
                                            for (int i = 0; i < total_cells.Count(); i++)
                                            {
                                                string row = "";
                                                rowinfo.TryGetValue(total_cells[i], out row);

                                                if (total_cells.Count() == 2)
                                                {
                                                    if (formula.EndsWith("("))
                                                        formula = formula + col[0] + row;
                                                    else
                                                        formula = formula + "," + col[0] + row;
                                                }
                                                else
                                                {
                                                    if (formula.EndsWith("("))
                                                    {
                                                        formula = formula + col[0] + row;
                                                        rowinfo.TryGetValue(total_cells[total_cells.Count() - 1], out row);
                                                        formula = formula + ":" + col[0] + row;
                                                    }
                                                }
                                            }
                                            formula = formula + ")";
                                            fc.formula = formula;
                                            fc.col = col[0];
                                            fc.row = headername_location[1];
                                            fc.total_name = HeaderFieldName;

                                            Boolean found = false;
                                            foreach (Formula_Class f in formulas)
                                            {
                                                double double_result = 0;
                                                if (Double.TryParse(f.formula.Substring(6, 1), out double_result) == true)
                                                {
                                                    if (f.formula.Substring(5, 1) == col[0] && f.total_name == HeaderFieldName)
                                                        found = true;
                                                }
                                                else
                                                {
                                                    if (f.formula.Substring(5, 2) == col[0] && f.total_name == HeaderFieldName)
                                                        found = true;
                                                }
                                            }
                                            if (found == false)
                                            {
                                                try
                                                {
                                                    formulas.Add(fc);
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine(ex.Message.ToString());
                                                }
                                                if (!ForTotalConstructionCostCol.ContainsValue(col[0]))
                                                    ForTotalConstructionCostCol.Add(ForTotalConstructionCostCol.Count.ToString(), col[0]);
                                            }
                                        }
                                        Console.WriteLine("");
                                        /*
                                                    

                                                    
                                                             Boolean found = false;
                                                             foreach (Formula_Class f in formulas)
                                                             {
                                                                 double double_result = 0;
                                                                 if (Double.TryParse(f.formula.Substring(6, 1), out double_result) == true)
                                                                 {
                                                                     if (f.formula.Substring(5, 1) == col[0] && f.total_name == HeaderFieldName)
                                                                         found = true;
                                                                 }
                                                                 else
                                                                 {
                                                                     if (f.formula.Substring(5, 2) == col[0] && f.total_name == HeaderFieldName)
                                                                         found = true;
                                                                 }
                                                             }
                                                             if (found == false)
                                                             {
                                                                 formulas.Add(fc);
                                                                 if (!ForTotalConstructionCostCol.ContainsValue(col[0]))
                                                                     ForTotalConstructionCostCol.Add(ForTotalConstructionCostCol.Count.ToString(), col[0]);
                                                             }
                                                         }
                                                         if (total_cells[total_cells.Count() - 1] == seq)
                                                         {
                                                             foreach (Formula_Class fc in formulas)
                                                             {
                                                                 double double_result = 0;
                                                                 if (Double.TryParse(fc.formula.Substring(6, 1), out double_result) == true)
                                                                 {
                                                                     if (fc.formula.Substring(5, 1) == col[0] && fc.total_name == HeaderFieldName)
                                                                         if (!fc.formula.EndsWith(")"))
                                                                             fc.formula = fc.formula + ":" + col[0] + row + ")";
                                                                 }
                                                                 else
                                                                 {
                                                                     if (fc.formula.Substring(5, 2) == col[0] && fc.total_name == HeaderFieldName)
                                                                         if (!fc.formula.EndsWith(")"))
                                                                             fc.formula = fc.formula + ":" + col[0] + row + ")";
                                                                 }
                                                             }
                                                         }*/

                                    }
                                }
                            }
                        }
                        Console.WriteLine("");

                        // Only deal with Total direct cost and Total indirect cost
                        Dictionary<string, Dictionary<string, string[]>> TotalCost = ReadTotalCostFromDB(Year_End, Budget_Project_On_Hand_Status);
                        foreach (string HeaderFieldName in ReportItem_Detail_Total.Keys)
                        {
                            string[] headername_location = new string[4];
                            ReportItem_Detail_Total.TryGetValue(HeaderFieldName, out headername_location);
                            if (headername_location != null)
                            {
                                if (!string.IsNullOrEmpty(headername_location[0]) && !string.IsNullOrEmpty(headername_location[1]))
                                {
                                    oRng = oSheet.get_Range(headername_location[0] + headername_location[1], headername_location[0] + headername_location[1]);
                                    string value = "";
                                    oRng.set_Value(Missing, HeaderFieldName);
                                }
                                if (headername_location.Count() > 2)
                                {
                                    if (!string.IsNullOrEmpty(headername_location[2]))
                                    {
                                        string[] total_cells = headername_location[2].Split('+');
                                        for (int i = 0; i < total_cells.Count(); i++)
                                        {
                                            total_cells[i] = total_cells[i].Replace("[", "").Replace("]", "");
                                        }
                                        string formula = "";
                                        once = false;
                                        foreach (string Budget_CashFlow_ReportItem_Name in TotalCost.Keys)
                                        {
                                            Dictionary<string, string[]> data;
                                            TotalCost.TryGetValue(Budget_CashFlow_ReportItem_Name, out data);

                                            foreach (string column_name in data.Keys)
                                            {
                                                string[] location_and_value;
                                                data.TryGetValue(column_name, out location_and_value);
                                                string row = location_and_value[0];
                                                string value = location_and_value[1];
                                                string type = location_and_value[2];
                                                string yrend_period = location_and_value[3];
                                                string seq = location_and_value[4];

                                                if (HeaderFieldName == "Total Construction Costs")
                                                    Console.WriteLine("");
                                                if (headername_location[3] == type)
                                                {
                                                    if (HeaderFieldName == "Total Construction Costs")
                                                        Console.WriteLine("");
                                                    if (once == false)
                                                    {
                                                        YearEnd = yrend_period;
                                                        DetailCol = ReadDetailColFromDB(type, yrend_period);
                                                        once = true;
                                                    }
                                                    string col = "";
                                                    if (DetailCol.Count > 0)
                                                        DetailCol.TryGetValue(column_name, out col);

                                                    if (col == "F")
                                                        Console.WriteLine("");
                                                    if (!string.IsNullOrEmpty(col))
                                                    {
                                                        if (total_cells[0] == seq)
                                                        {
                                                            string old_col = col;
                                                            foreach (string k in ForTotalConstructionCostCol.Keys)
                                                            {

                                                                Formula_Class fc = new Formula_Class();

                                                                ForTotalConstructionCostCol.TryGetValue(k, out col);
                                                                formula = "=Sum(" + col + row;
                                                                fc.formula = formula;
                                                                fc.col = col;
                                                                fc.row = headername_location[1];
                                                                fc.total_name = HeaderFieldName;


                                                                Boolean found = false;
                                                                foreach (Formula_Class f in formulas)
                                                                {
                                                                    double double_result = 0;
                                                                    if (Double.TryParse(f.formula.Substring(6, 1), out double_result) == true)
                                                                    {
                                                                        if (f.formula.Substring(5, 1) == col && f.total_name == HeaderFieldName)
                                                                            found = true;
                                                                    }
                                                                    else
                                                                    {
                                                                        if (f.formula.Substring(5, 2) == col && f.total_name == HeaderFieldName)
                                                                            found = true;
                                                                    }
                                                                }
                                                                if (found == false)
                                                                {
                                                                    try
                                                                    {
                                                                        formulas.Add(fc);
                                                                    }
                                                                    catch (Exception ex)
                                                                    {
                                                                        Console.WriteLine(ex.Message.ToString());
                                                                    }
                                                                }
                                                            }
                                                            col = old_col;
                                                        }
                                                        if (total_cells[total_cells.Count() - 1] == seq)
                                                        {
                                                            foreach (string k in ForTotalConstructionCostCol.Keys)
                                                            {
                                                                ForTotalConstructionCostCol.TryGetValue(k, out col);
                                                                foreach (Formula_Class fc in formulas)
                                                                {
                                                                    double double_result = 0;
                                                                    if (Double.TryParse(fc.formula.Substring(6, 1), out double_result) == true)
                                                                    {
                                                                        if (fc.formula.Substring(5, 1) == col && fc.total_name == HeaderFieldName)
                                                                            if (!fc.formula.EndsWith(")"))
                                                                                fc.formula = fc.formula + "," + col + row + ")"; // differece at ","
                                                                    }
                                                                    else
                                                                    {
                                                                        if (fc.formula.Substring(5, 2) == col && fc.total_name == HeaderFieldName)
                                                                            if (!fc.formula.EndsWith(")"))
                                                                                fc.formula = fc.formula + "," + col + row + ")"; // differece at ","
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        Console.WriteLine("");
                        foreach (Formula_Class fc in formulas)
                        {
                            oRng = oSheet.get_Range(fc.col + fc.row, fc.col + fc.row);
                            oRng.set_Value(Missing, System.String.Format("0", "##,###,###,##0.00"));
                            oRng.NumberFormatLocal = "#,##0.00_);(#,##0.00)";
                            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                            if (fc.formula.Contains(":") || fc.formula.Contains(","))
                            {
                                try
                                {
                                    oRng.set_Value(Missing, fc.formula);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("");
                                }
                            }
                            else
                            {
                                try
                                {
                                    oRng.set_Value(Missing, fc.formula + ":" + fc.formula.Replace("=Sum(", "") + ")");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("");
                                }
                            }
                        }

                        Dictionary<string, string[]> HeaderDetail = ReadHeaderDetailFromDB(Year_End, Budget_Project_On_Hand_Status);
                        foreach (string HeaderFieldName in HeaderDetail.Keys)
                        {
                            string[] headername_location = new string[3];
                            HeaderDetail.TryGetValue(HeaderFieldName, out headername_location);
                            if (headername_location != null)
                            {
                                if (!string.IsNullOrEmpty(headername_location[0]) && !string.IsNullOrEmpty(headername_location[1]))
                                {
                                    oRng = oSheet.get_Range(headername_location[1] + headername_location[0], headername_location[1] + headername_location[0]);
                                    string value = "";
                                    oRng.set_Value(Missing, headername_location[2]);
                                }
                            }
                        }
                        oRng = oSheet.get_Range("D3", "D3");// Hard code
                        oRng.set_Value(Missing, Budget_Project_On_Hand_Status);
                        oRng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                        string Title_Year_End = "";
                        Title_Year_End = day.ToString() + " " + Convert.ToDateTime(YrEnd_Date).ToString("MMM") + " " + year.ToString();

                        oRng = oSheet.get_Range("A1", "A1"); // Hard code
                        oRng.set_Value(Missing, "Project Budget For The Year Ended " + Title_Year_End);
                        if (Detail.Count == 0)
                            MessageBox.Show("No data");

                        oRng = oSheet.get_Range("A1", "Z65536");
                        oRng.EntireColumn.AutoFit();

                        oSheet.Columns.HorizontalAlignment = 1;
                        oSheet.Columns.VerticalAlignment = -4160;
                        oSheet.Columns.Orientation = 0;
                        oSheet.Columns.AddIndent = false;
                        oSheet.Columns.IndentLevel = 0;
                        oSheet.Columns.ShrinkToFit = false;
                        oSheet.Columns.ReadingOrder = -5002;
                        oSheet.Columns.MergeCells = false;
                        oSheet.Columns.Font.Name = "Arial";
                        oSheet.Columns.Font.FontStyle = "Normal";
                        oSheet.Columns.Font.Size = 10;
                        oSheet.Columns.Font.Strikethrough = false;
                        oSheet.Columns.Font.Superscript = false;
                        oSheet.Columns.Font.Subscript = false;
                        oSheet.Columns.Font.OutlineFont = false;
                        oSheet.Columns.Font.Shadow = false;
                        oSheet.Columns.Font.Underline = -4142;
                        oSheet.Columns.Font.ColorIndex = -4105;
                        if (oSheet.Cells[3, 1] == "Project Name")
                        {
                            oRng = oSheet.get_Range("A1", "A65536");
                            oRng.ColumnWidth = 114;
                        }

                        oRng = oSheet.get_Range("A1", "A65536");
                        //oXL.Visible = true;

                        for (int i = 0; i < 10; i++)
                            progressBar1.PerformStep();
                    }
                    
                    //string fname = "D:\\Data\\My Documents\\Visual Studio 2008\\WebSites\\Cheques\\Reports\\ChequeDataEnquiry.xls";

                    /*
                    string fname = Application.StartupPath + "\\Excel_" + DateTime.Now.ToString().Replace(":", "_").Replace("/", "-") + ".xls";


                    if (System.IO.File.Exists(fname))
                    {
                        System.IO.File.Delete(fname);
                        oWB.SaveCopyAs(fname);

                        //MessageBox.Show("Export Success");
                    }
                    else
                    {
                        oWB.SaveCopyAs(fname);
                        //MessageBox.Show("Export Success");
                    }

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
                    */
                    oXL.Visible = true;
                    //MessageBox.Show(fname);
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
                //this.Close();
                Application.Exit();
            }
        }
        static Dictionary<string, string> ReadDetailColFromDB(string cashflow_type, string yrend_calperiod)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    //command.CommandText = "select DataField_Name, Col_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','CumTo_ForPeriod','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "'";

                    string sql = "";
                    sql += "select DataField_Name, Col_No from Budget_CashFlow_Excel_ReportItem_RowCol ";
                    sql += "where Budget_CashFlow_Type = '" + cashflow_type + "' ";
                    sql += "and YrEnd_CalPeriod = '"+yrend_calperiod+"' ";
                    sql += "and DataField_Name not in ('[START_ROW_NO]','Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code')";
                    //sql += "and DataField_Name not in ('[START_ROW_NO]','Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code') and Heading_Type = 'D'";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["DataField_Name"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                            {
                                result.Add(dr["DataField_Name"].ToString(), dr["Col_No"].ToString());
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
        static Dictionary<string, string[]> ReadDetailColTotalFromDB(string cashflow_type, string yrend_calperiod)
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

                    string sql = "";
                    sql += "select Budget_CashFlow_Excel_ReportItem_RowCol_ID, Col_No, Total_Formula, Heading_Type from Budget_CashFlow_Excel_ReportItem_RowCol ";
                    sql += "where Budget_CashFlow_Type = '" + cashflow_type + "' ";
                    sql += "and YrEnd_CalPeriod = '" + yrend_calperiod + "' ";
                    sql += "and DataField_Name not in ('[START_ROW_NO]','Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code')";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Budget_CashFlow_Excel_ReportItem_RowCol_ID"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                            {
                                string[] data = new string[3];
                                data[0] = dr["Col_No"].ToString();
                                data[1] = dr["Total_Formula"].ToString();
                                data[2] = dr["Heading_Type"].ToString();
                                result.Add(dr["Budget_CashFlow_Excel_ReportItem_RowCol_ID"].ToString(), data);
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
        static Dictionary<string, string[]> ReadDetailColTotalTotalFromDB(string cashflow_type, string yrend_calperiod)
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

                    string sql = "";
                    sql += "select * from Budget_CashFlow_Excel_ReportItem_RowCol ";
                    sql += " where YrEnd_CalPeriod = '" + yrend_calperiod + "'";
                    sql += " and Budget_CashFlow_Type = '" + cashflow_type + "'";
                    sql += " and Heading_Type = 'T'";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["DataField_Name"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                            {
                                string[] data = new string[3];
                                data[0] = dr["Col_No"].ToString();
                                data[1] = dr["Total_Formula"].ToString();
                                data[2] = dr["Heading_Type"].ToString();
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
        static Dictionary<string, string> ReadTotalTotalRowInfoFromDB(string year_end, string Budget_Project_On_Hand_Status)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    //command.CommandText = "select DataField_Name, Col_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','CumTo_ForPeriod','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "'";

                    string sql = "";
                    if (string.IsNullOrEmpty(year_end))
                    {
                        sql += "select * from dbo.Budget_CashFlow_Excel_ReportItem_Detail ";
                        sql += " where Heading_Type in ('T') ";
                        sql += " and Budget_Project_On_Hand_Status='" + Budget_Project_On_Hand_Status + "' ";
                        sql += " and YrEnd_CalPeriod in ( ";
                        sql += " Select top 1 YrEnd_CalPeriod from dbo.Budget_CashFlow_Excel_ReportItem_Detail ";
                        sql += " where  ";
                        sql += " Budget_Project_On_Hand_Status='" + Budget_Project_On_Hand_Status + "' order by YrEnd_CalPeriod desc)";
                    }
                    else
                    {
                        sql += "select * from dbo.Budget_CashFlow_Excel_ReportItem_Detail ";
                        sql += " where ";
                        sql += " Budget_Project_On_Hand_Status='" + Budget_Project_On_Hand_Status + "'";
                        sql += " and YrEnd_CalPeriod='" + year_end + "'";
                    }

                    int counter = 0;
                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Row_No"] != System.DBNull.Value)
                        {
                            if (dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"] != System.DBNull.Value)
                            {
                                result.Add(dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString(), dr["Row_No"].ToString());
                                counter += 1;
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
        static Dictionary<string, Dictionary<string, string[]>> ReadTotalCostFromDB(string year_end, string Budget_Project_On_Hand_Status)
        {
            Dictionary<string, Dictionary<string, string[]>> result = new Dictionary<string, Dictionary<string, string[]>>();
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    //command.CommandText = "select DataField_Name, Col_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','CumTo_ForPeriod','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "'";

                    string sql = "";
                    if (string.IsNullOrEmpty(year_end))
                    {
                        sql += "select * from dbo.Budget_CashFlow_Excel_ReportItem_Detail ";
                        sql += " where Heading_Type in ('T') ";
                        sql += " and Budget_Project_On_Hand_Status='" + Budget_Project_On_Hand_Status + "' ";
                        sql += " and YrEnd_CalPeriod in ( ";
                        sql += " Select top 1 YrEnd_CalPeriod from dbo.Budget_CashFlow_Excel_ReportItem_Detail ";
                        sql += " where Heading_Type in ('T') ";
                        sql += " and Budget_Project_On_Hand_Status='" + Budget_Project_On_Hand_Status + "' order by YrEnd_CalPeriod desc)";
                    }
                    else
                    {
                        sql += "select * from dbo.Budget_CashFlow_Excel_ReportItem_Detail ";
                        sql += " where Heading_Type in ('T') ";
                        sql += " and Budget_Project_On_Hand_Status='" + Budget_Project_On_Hand_Status + "'";
                        sql += " and YrEnd_CalPeriod='" + year_end + "'";
                    }

                    int counter = 0;
                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Row_No"] != System.DBNull.Value)
                        {
                            if (dr["Budget_CashFlow_ReportItem_Name"] != System.DBNull.Value && dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"] != System.DBNull.Value)
                            {
                                Dictionary<string, string[]> column_data = new Dictionary<string, string[]>();
                                string[] data1 = new string[5];
                                data1[0] = dr["Row_No"].ToString();
                                data1[1] = "0";
                                //if (dr["CumTo_ForPeriod"] == System.DBNull.Value)
                                    //data1[1] = "0";
                                //else
                                    //data1[1] = dr["CumTo_ForPeriod"].ToString();
                                data1[2] = dr["Budget_CashFlow_Type"].ToString();
                                data1[3] = dr["YrEnd_CalPeriod"].ToString();
                                data1[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("CumTo_ForPeriod", data1);

                                result.Add(dr["Budget_CashFlow_ReportItem_Name"].ToString() + "$" + counter.ToString(), column_data);
                                counter += 1;
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
        static Dictionary<string, Dictionary<string, string[]>> ReadDetailRowFromDB(string project_code, string YrEnd_CalPeriod)
        {
            Dictionary<string, Dictionary<string,string[]>> result = new Dictionary<string, Dictionary<string,string[]>>();
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    //command.CommandText = "select DataField_Name, Col_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','CumTo_ForPeriod','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "'";

                    string sql = "";
                    sql += "select c.YrEnd_CalPeriod,c.Row_No,c.Col_No, c.Heading_Type,c.Budget_CashFlow_Excel_ReportItem_Detail_ID, c.Budget_Project_On_Hand_Status,c.Budget_CashFlow_Type,* from dbo.Budget_CashFlow_Detail a, ";
                    sql += "(select Last_Upload_Budget_CashFlow_Worksheet_ID ";
                    sql += "from dbo.Budget_CashFlow_Project where Budget_Project_No = '" + project_code + "' and YrEnd_CalPeriod='" + YrEnd_CalPeriod + "') as b, ";
                    sql += "(select Budget_Project_On_Hand_Status, Budget_CashFlow_Worksheet_ID ";
                    sql += "from dbo.Budget_CashFlow_Worksheet ";
                    sql += "where Budget_Project_No = '" + project_code + "' ) as d, ";
                    sql += "Budget_CashFlow_Excel_ReportItem_Detail c ";
                    sql += "where b.Last_Upload_Budget_CashFlow_Worksheet_ID = a.Budget_CashFlow_Worksheet_ID ";
                    sql += "and d.Budget_CashFlow_Worksheet_ID = b.Last_Upload_Budget_CashFlow_Worksheet_ID ";
                    sql += "and a.Budget_CashFlow_ReportItem_ID = c.Budget_CashFlow_ReportItem_ID ";
                    sql += "and d.Budget_Project_On_Hand_Status = c.Budget_Project_On_Hand_Status";

                    int counter = 0;
                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Row_No"] != System.DBNull.Value)
                        {
                            if (dr["Budget_CashFlow_ReportItem_Name"] != System.DBNull.Value && dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"] != System.DBNull.Value)
                            {
                                Dictionary<string, string[]> column_data = new Dictionary<string, string[]>();
                                string[] data1 = new string[5];
                                data1[0] = dr["Row_No"].ToString();
                                if( dr["CumTo_ForPeriod"] == System.DBNull.Value)
                                    data1[1] = "0";
                                else
                                    data1[1] = dr["CumTo_ForPeriod"].ToString();
                                data1[2] = dr["Budget_CashFlow_Type"].ToString();
                                data1[3] = dr["YrEnd_CalPeriod"].ToString();
                                data1[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("CumTo_ForPeriod",data1);

                                string[] data2 = new string[5];
                                data2[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod1"] == System.DBNull.Value)
                                    data2[1] = "0";
                                else
                                    data2[1] = dr["ForPeriod1"].ToString();
                                data2[2] = dr["Budget_CashFlow_Type"].ToString();
                                data2[3] = dr["YrEnd_CalPeriod"].ToString();
                                data2[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod1", data2);

                                string[] data3 = new string[5];
                                data3[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod2"] == System.DBNull.Value)
                                    data3[1] = "0";
                                else
                                    data3[1] = dr["ForPeriod2"].ToString();
                                data3[2] = dr["Budget_CashFlow_Type"].ToString();
                                data3[3] = dr["YrEnd_CalPeriod"].ToString();
                                data3[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod2", data3);

                                string[] data4 = new string[5];
                                data4[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod3"] == System.DBNull.Value)
                                    data4[1] = "0";
                                else
                                    data4[1] = dr["ForPeriod3"].ToString();
                                data4[2] = dr["Budget_CashFlow_Type"].ToString();
                                data4[3] = dr["YrEnd_CalPeriod"].ToString();
                                data4[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod3", data4);

                                string[] data5 = new string[5];
                                data5[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod4"] == System.DBNull.Value)
                                    data5[1] = "0";
                                else
                                    data5[1] = dr["ForPeriod4"].ToString();
                                data5[2] = dr["Budget_CashFlow_Type"].ToString();
                                data5[3] = dr["YrEnd_CalPeriod"].ToString();
                                data5[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod4", data5);

                                string[] data6 = new string[5];
                                data6[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod5"] == System.DBNull.Value)
                                    data6[1] = "0";
                                else
                                    data6[1] = dr["ForPeriod5"].ToString();
                                data6[2] = dr["Budget_CashFlow_Type"].ToString();
                                data6[3] = dr["YrEnd_CalPeriod"].ToString();
                                data6[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod5", data6);

                                string[] data7 = new string[5];
                                data7[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod6"] == System.DBNull.Value)
                                    data7[1] = "0";
                                else
                                    data7[1] = dr["ForPeriod6"].ToString();
                                data7[2] = dr["Budget_CashFlow_Type"].ToString();
                                data7[3] = dr["YrEnd_CalPeriod"].ToString();
                                data7[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod6", data7);

                                string[] data8 = new string[5];
                                data8[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod7"] == System.DBNull.Value)
                                    data8[1] = "0";
                                else
                                    data8[1] = dr["ForPeriod7"].ToString();
                                data8[2] = dr["Budget_CashFlow_Type"].ToString();
                                data8[3] = dr["YrEnd_CalPeriod"].ToString();
                                data8[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod7", data8);

                                string[] data9 = new string[5];
                                data9[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod8"] == System.DBNull.Value)
                                    data9[1] = "0";
                                else
                                    data9[1] = dr["ForPeriod8"].ToString();
                                data9[2] = dr["Budget_CashFlow_Type"].ToString();
                                data9[3] = dr["YrEnd_CalPeriod"].ToString();
                                data9[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod8", data9);

                                string[] data10 = new string[5];
                                data10[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod9"] == System.DBNull.Value)
                                    data10[1] = "0";
                                else
                                    data10[1] = dr["ForPeriod9"].ToString();
                                data10[2] = dr["Budget_CashFlow_Type"].ToString();
                                data10[3] = dr["YrEnd_CalPeriod"].ToString();
                                data10[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod9", data10);

                                string[] data11 = new string[5];
                                data11[0] = dr["Row_No"].ToString();
                                if(dr["ForPeriod10"] == System.DBNull.Value)
                                    data11[1] = "0";
                                else
                                    data11[1] = dr["ForPeriod10"].ToString();
                                data11[2] = dr["Budget_CashFlow_Type"].ToString();
                                data11[3] = dr["YrEnd_CalPeriod"].ToString();
                                data11[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod10", data11);

                                string[] data12 = new string[5];
                                data12[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod11"] == System.DBNull.Value)
                                    data12[1] = "0";
                                else
                                    data12[1] = dr["ForPeriod11"].ToString();
                                data12[2] = dr["Budget_CashFlow_Type"].ToString();
                                data12[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod11", data12);

                                string[] data13 = new string[5];
                                data13[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod12"] == System.DBNull.Value)
                                    data13[1] = "0";
                                else
                                    data13[1] = dr["ForPeriod12"].ToString();
                                data13[2] = dr["Budget_CashFlow_Type"].ToString();
                                data13[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod12", data13);

                                string[] data14 = new string[5];
                                data14[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod13"] == System.DBNull.Value)
                                    data14[1] = "0";
                                else
                                    data14[1] = dr["ForPeriod13"].ToString();
                                data14[2] = dr["Budget_CashFlow_Type"].ToString();
                                data14[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod13", data14);

                                string[] data15 = new string[5];
                                data15[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod14"] == System.DBNull.Value)
                                    data15[1] = "0";
                                else
                                    data15[1] = dr["ForPeriod14"].ToString();
                                data15[2] = dr["Budget_CashFlow_Type"].ToString();
                                data15[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod14", data15);

                                string[] data16 = new string[5];
                                data16[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod15"] == System.DBNull.Value)
                                    data16[1] = "0";
                                else
                                    data16[1] = dr["ForPeriod15"].ToString();
                                data16[2] = dr["Budget_CashFlow_Type"].ToString();
                                data16[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod15", data16);

                                string[] data17 = new string[5];
                                data17[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod16"] == System.DBNull.Value)
                                    data17[1] = "0";
                                else
                                    data17[1] = dr["ForPeriod16"].ToString();
                                data17[2] = dr["Budget_CashFlow_Type"].ToString();
                                data17[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod16", data17);

                                string[] data18 = new string[5];
                                data18[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod17"] == System.DBNull.Value)
                                    data18[1] = "0";
                                else
                                    data18[1] = dr["ForPeriod17"].ToString();
                                data18[2] = dr["Budget_CashFlow_Type"].ToString();
                                data18[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod17", data18);

                                string[] data19 = new string[5];
                                data19[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod18"]== System.DBNull.Value)
                                    data19[1] = "0";
                                else
                                    data19[1] = dr["ForPeriod18"].ToString();
                                data19[2] = dr["Budget_CashFlow_Type"].ToString();
                                data19[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod18", data19);

                                string[] data20 = new string[5];
                                data20[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod19"]== System.DBNull.Value)
                                    data20[1] = "0";
                                else
                                    data20[1] = dr["ForPeriod19"].ToString();
                                data20[2] = dr["Budget_CashFlow_Type"].ToString();
                                data20[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod19", data20);

                                string[] data21 = new string[5];
                                data21[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod20"] == System.DBNull.Value)
                                    data21[1] = "0";
                                else
                                    data21[1] = dr["ForPeriod20"].ToString();
                                data21[2] = dr["Budget_CashFlow_Type"].ToString();
                                data21[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod20", data21);

                                string[] data22 = new string[5];
                                data22[0] = dr["Row_No"].ToString();
                                data22[1] = dr["ForPeriod21"].ToString();
                                data22[2] = dr["Budget_CashFlow_Type"].ToString();
                                data22[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod21", data22);

                                string[] data23 = new string[5];
                                data23[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod22"]== System.DBNull.Value)
                                    data23[1] = "0";
                                else
                                    data23[1] = dr["ForPeriod22"].ToString();
                                data23[2] = dr["Budget_CashFlow_Type"].ToString();
                                data23[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod22", data23);

                                string[] data24 = new string[5];
                                data24[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod23"] == System.DBNull.Value)
                                    data24[1] = "0";
                                else
                                    data24[1] = dr["ForPeriod23"].ToString();
                                data24[2] = dr["Budget_CashFlow_Type"].ToString();
                                data24[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod23", data24);

                                string[] data25 = new string[5];
                                data25[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod24"] == System.DBNull.Value)
                                    data25[1] = "0";
                                else
                                    data25[1] = dr["ForPeriod24"].ToString();
                                data25[2] = dr["Budget_CashFlow_Type"].ToString();
                                data25[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod24", data25);

                                string[] data26 = new string[5];
                                data26[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriod_After"]== System.DBNull.Value)
                                    data26[1] = "0";
                                else
                                    data26[1] = dr["ForPeriod_After"].ToString();
                                data26[2] = dr["Budget_CashFlow_Type"].ToString();
                                data26[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriod_After", data26);

                                string[] data27 = new string[5];
                                data27[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriodC9"]== System.DBNull.Value)
                                    data27[1] = "0";
                                else
                                    data27[1] = dr["ForPeriodC9"].ToString();
                                data27[2] = dr["Budget_CashFlow_Type"].ToString();
                                data27[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriodC9", data27);

                                string[] data28 = new string[5];
                                data28[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriodC10"]== System.DBNull.Value)
                                    data28[1] = "0";
                                else
                                    data28[1] = dr["ForPeriodC10"].ToString();
                                data28[2] = dr["Budget_CashFlow_Type"].ToString();
                                data28[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriodC10", data28);

                                string[] data29 = new string[5];
                                data29[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriodC11"]== System.DBNull.Value)
                                    data29[1] = "0";
                                else
                                    data29[1] = dr["ForPeriodC11"].ToString();
                                data29[2] = dr["Budget_CashFlow_Type"].ToString();
                                data29[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriodC11", data29);

                                string[] data30 = new string[5];
                                data30[0] = dr["Row_No"].ToString();
                                if( dr["ForPeriodC12"] == System.DBNull.Value)
                                    data30[1] = "0";
                                else
                                    data30[1] = dr["ForPeriodC12"].ToString();
                                data30[2] = dr["Budget_CashFlow_Type"].ToString();
                                data30[4] = dr["Budget_CashFlow_Excel_ReportItem_Detail_ID"].ToString();
                                column_data.Add("ForPeriodC12", data30);

                                result.Add(dr["Budget_CashFlow_ReportItem_Name"].ToString() +"$"+ counter.ToString(), column_data);
                                counter += 1;
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
        }/*
        static Dictionary<string, string[]> ReadHeaderRowColFromDB(string yrend_calperiod, string Budget_Project_On_Hand_Status)
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

                    string sql = "";
                    sql += "select DataField_Name, Col_No, Row_No from Budget_CashFlow_Excel_HeaderItem_RowCol ";
                    sql += "where YrEnd_CalPeriod = '" + yrend_calperiod + "' ";
                    sql += "Budget_Project_On_Hand_Status = '" + Budget_Project_On_Hand_Status + "' ";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["DataField_Name"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                                if (dr["Row_No"] != System.DBNull.Value)
                                {
                                    string[] data = new string[2];
                                    data[0] = dr["Row_No"].ToString();
                                    data[1] = dr["Col_No"].ToString();
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
        }*/
        static Dictionary<string, string[]> ReadHeaderDetailFromDB(string yrend_calperiod, string Budget_Project_On_Hand_Status)
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

                    string sql = "";
                    sql += "select Budget_CashFlow_Excel_HeaderItem_Detail_ID, Upload_Excel_Field_Name, Col_No, Row_No from Budget_CashFlow_Excel_HeaderItem_Detail ";
                    sql += "where YrEnd_CalPeriod = '" + yrend_calperiod + "' ";
                    sql += " and Budget_Project_On_Hand_Status = '" + Budget_Project_On_Hand_Status + "'";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    int count = 0;
                    while (dr.Read())
                    {
                        if (dr["Upload_Excel_Field_Name"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                                if (dr["Row_No"] != System.DBNull.Value)
                                {
                                    string[] data = new string[3];
                                    data[0] = dr["Row_No"].ToString();
                                    data[1] = dr["Col_No"].ToString();
                                    data[2] = dr["Upload_Excel_Field_Name"].ToString();

                                    result.Add(dr["Budget_CashFlow_Excel_HeaderItem_Detail_ID"].ToString(), data);
                                    count = count + 1;
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
        static Dictionary<string, string[]> ReadHeaderRowColFromDB(string yrend_calperiod, string Budget_Project_On_Hand_Status)
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

                    string sql = "";
                    sql += "select DataField_Name, Col_No, Row_No, DataField_Type from Budget_CashFlow_Excel_HeaderItem_RowCol ";
                    sql += "where YrEnd_CalPeriod = '"+yrend_calperiod+"' ";
                    sql += " and Budget_Project_On_Hand_Status = '" + Budget_Project_On_Hand_Status + "'";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["DataField_Name"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                                if (dr["Row_No"] != System.DBNull.Value)
                                {
                                    string[] data = new string[3];
                                    data[0] = dr["Row_No"].ToString();
                                    data[1] = dr["Col_No"].ToString();
                                    if (dr["DataField_Type"] != System.DBNull.Value)
                                        data[2] = dr["DataField_Type"].ToString();
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
        static int[] ReadHeaderRowColPeriodRowFromDB(string yrend_calperiod, string Budget_Project_On_Hand_Status)
        {
            int[] result = new int[2];
            for (int i = 0; i < result.Count(); i++)
                result[i] = 0;

            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    //command.CommandText = "select DataField_Name, Col_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','CumTo_ForPeriod','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "'";

                    string sql = "";
                    sql += "select Budget_CashFlow_Excel_HeaderItem_RowCol_ID, DataField_Name, Col_No, Row_No, DataField_Type, Excel_Cell_Formula from Budget_CashFlow_Excel_HeaderItem_RowCol ";
                    sql += "where YrEnd_CalPeriod = '" + yrend_calperiod + "' ";
                    sql += " and Budget_Project_On_Hand_Status = '" + Budget_Project_On_Hand_Status + "' and Heading_Type='F' and DataField_Name='[PERIOD_ROW]'";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    int count = 0;
                    while (dr.Read())
                    {
                        if (dr["Row_No"] != System.DBNull.Value)
                        {
                            result[count] = Convert.ToInt32(dr["Row_No"].ToString());
                            count = count + 1;
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
        static Dictionary<string, string[]> ReadHeaderRowColFormulaFromDB(string yrend_calperiod, string Budget_Project_On_Hand_Status)
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

                    string sql = "";
                    sql += "select Budget_CashFlow_Excel_HeaderItem_RowCol_ID, DataField_Name, Col_No, Row_No, DataField_Type, Excel_Cell_Formula from Budget_CashFlow_Excel_HeaderItem_RowCol ";
                    sql += "where YrEnd_CalPeriod = '" + yrend_calperiod + "' ";
                    sql += " and Budget_Project_On_Hand_Status = '" + Budget_Project_On_Hand_Status + "' and Heading_Type='F'";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Budget_CashFlow_Excel_HeaderItem_RowCol_ID"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                                if (dr["Row_No"] != System.DBNull.Value)
                                {
                                    string[] data = new string[3];
                                    data[0] = dr["Row_No"].ToString();
                                    data[1] = dr["Col_No"].ToString();
                                    if (dr["Excel_Cell_Formula"] != System.DBNull.Value)
                                        data[2] = dr["Excel_Cell_Formula"].ToString();
                                    result.Add(dr["Budget_CashFlow_Excel_HeaderItem_RowCol_ID"].ToString(), data);
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
        static Dictionary<string, string[]> ReadHeaderRowColFromDB(string YrEndDate, string Budget_Project_On_Hand_Status)
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

                    string sql = "";
                    sql += "select * from dbo.Budget_CashFlow_Worksheet a, ";
                    sql += "(select Last_Upload_Budget_CashFlow_Worksheet_ID ";
                    sql += "from dbo.Budget_CashFlow_Project where Budget_Project_No = '" + project_code + "') as b, ";
                    sql += "(select Budget_Project_On_Hand_Status, Budget_CashFlow_Worksheet_ID ";
                    sql += "from dbo.Budget_CashFlow_Worksheet ";
                    sql += "where Budget_Project_No = '" + project_code + "' ) as d ";
                    sql += "where b.Last_Upload_Budget_CashFlow_Worksheet_ID = a.Budget_CashFlow_Worksheet_ID ";
                    sql += "and d.Budget_CashFlow_Worksheet_ID = b.Last_Upload_Budget_CashFlow_Worksheet_ID";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Row_No"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                            {
                                string[] location = new string[2];
                                location[0] = dr["Row_No"].ToString();
                                location[1] = dr["Col_No"].ToString();
                                result.Add("Budget_Project_On_Hand_Status", location);
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
        }*/
        static Dictionary<string, string[]> ReadReportItem_Detail_Total(string Budget_Project_On_Hand_Status)
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

                    string sql = "";
                    sql += "select * from dbo.Budget_CashFlow_Excel_ReportItem_Detail ";
                    sql += " where Heading_Type in ('T') ";
                    sql += " and Budget_Project_On_Hand_Status='" + Budget_Project_On_Hand_Status + "'";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Budget_CashFlow_ReportItem_Name"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                            {
                                if (dr["Row_No"] != System.DBNull.Value)
                                {
                                    string[] location = new string[4];
                                    location[0] = dr["Col_No"].ToString();
                                    location[1] = dr["Row_No"].ToString();
                                    if (dr["Total_Formula"] != System.DBNull.Value)
                                    {
                                        location[2] = dr["Total_Formula"].ToString();
                                    }
                                    if (dr["Budget_CashFlow_Type"] != System.DBNull.Value)
                                    {
                                        location[3] = dr["Budget_CashFlow_Type"].ToString();
                                    }
                                    result.Add(dr["Budget_CashFlow_ReportItem_Name"].ToString(), location);
                                }
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
        static Dictionary<string, string[]> ReadReportItem_Detail_Header(string Budget_Project_On_Hand_Status)
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

                    string sql = "";
                    sql += "select * from dbo.Budget_CashFlow_Excel_ReportItem_Detail ";
                    sql += " where Heading_Type in ('H') ";
                    sql += " and Budget_Project_On_Hand_Status='" + Budget_Project_On_Hand_Status + "'";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Budget_CashFlow_ReportItem_Name"] != System.DBNull.Value)
                        {
                            if (dr["Col_No"] != System.DBNull.Value)
                            {
                                if (dr["Row_No"] != System.DBNull.Value)
                                {
                                    string[] location = new string[2];
                                    location[0] = dr["Col_No"].ToString();
                                    location[1] = dr["Row_No"].ToString();
                                    result.Add(dr["Budget_CashFlow_ReportItem_Name"].ToString(), location);
                                }
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
        static Dictionary<string, string> ReadWorkSheetFromDB(string project_code)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    //command.CommandText = "select DataField_Name, Col_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','CumTo_ForPeriod','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "'";

                    string sql = "";
                    sql += "select * from dbo.Budget_CashFlow_Worksheet a, ";
                    sql += "(select Last_Upload_Budget_CashFlow_Worksheet_ID ";
                    sql += "from dbo.Budget_CashFlow_Project where Budget_Project_No = '" + project_code + "') as b, ";
                    sql += "(select Budget_Project_On_Hand_Status, Budget_CashFlow_Worksheet_ID ";
                    sql += "from dbo.Budget_CashFlow_Worksheet ";
                    sql += "where Budget_Project_No = '" + project_code + "' ) as d ";
                    sql += "where b.Last_Upload_Budget_CashFlow_Worksheet_ID = a.Budget_CashFlow_Worksheet_ID ";
                    sql += "and d.Budget_CashFlow_Worksheet_ID = b.Last_Upload_Budget_CashFlow_Worksheet_ID";

                    command.CommandText = sql;
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        /*
                        if (dr["Budget_CashFlow_Header_ID"] != System.DBNull.Value)
                        {
                             
                        }
                        if (dr["Budget_CashFlow_Worksheet_Name"] != System.DBNull.Value)
                        {
                             
                        }*/
                        if (dr["YrEnd_Date"] != System.DBNull.Value)
                        {
                            result.Add("YrEnd_Date", dr["YrEnd_Date"].ToString()); 
                        }
                        if (dr["Budget_Main_Contractor_ProjectNo"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Main_Contractor_ProjectNo", dr["Budget_Main_Contractor_ProjectNo"].ToString());
                        }
                        if (dr["Budget_Main_Contractor"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Main_Contractor", dr["Budget_Main_Contractor"].ToString());        
                        }
                        if (dr["Budget_Project_On_Hand_Status"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Project_On_Hand_Status", dr["Budget_Project_On_Hand_Status"].ToString());       
                        }
                        if (dr["Budget_Company_Name"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Company_Name", dr["Budget_Company_Name"].ToString());      
                        }
                        if (dr["Budget_Project_Name"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Project_Name", dr["Budget_Project_Name"].ToString());      
                        }
                        if (dr["Budget_Project_No"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Project_No", dr["Budget_Project_No"].ToString());     
                        }
                        if (dr["Budget_Customer_Name"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Customer_Name", dr["Budget_Customer_Name"].ToString());     
                        }
                        if (dr["Budget_Customer_Type"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Customer_Type", dr["Budget_Customer_Type"].ToString());    
                        }
                        if (dr["Budget_Hit_Rate"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Hit_Rate", dr["Budget_Hit_Rate"].ToString());   
                        }
                        if (dr["Budget_Nature_Of_Works"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Nature_Of_Works", dr["Budget_Nature_Of_Works"].ToString());  
                        }
                        if (dr["Budget_Percentage_Shared"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Percentage_Shared", dr["Budget_Percentage_Shared"].ToString()); 
                        }
                        if (dr["Budget_Estimated_Completed_Date"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Estimated_Completed_Date", dr["Budget_Estimated_Completed_Date"].ToString());
                        }
                        if (dr["Budget_Actual_Completed_Date"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Actual_Completed_Date", dr["Budget_Actual_Completed_Date"].ToString());     
                        }
                        if (dr["Budget_Date_of_award"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Date_of_award", dr["Budget_Date_of_award"].ToString());    
                        }
                        if (dr["Budget_Original_Contract_Sum"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Original_Contract_Sum", dr["Budget_Original_Contract_Sum"].ToString());   
                        }
                        if (dr["Budget_Estimated_Final_Contract_Sum"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Estimated_Final_Contract_Sum", dr["Budget_Estimated_Final_Contract_Sum"].ToString());  
                        }
                        if (dr["Budget_Estimated_Final_Costs"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Estimated_Final_Costs", dr["Budget_Estimated_Final_Costs"].ToString()); 
                        }
                        if (dr["Budget_Estimated_Final_Margin"] != System.DBNull.Value)
                        {
                            result.Add("Budget_Estimated_Final_Margin", dr["Budget_Estimated_Final_Margin"].ToString());
                        }
                        if (dr["CumTo_ForPeriod_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("CumTo_ForPeriod_YrMth", dr["CumTo_ForPeriod_YrMth"].ToString());                       
                        }
                        if (dr["ForPeriodC9_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriodC9_YrMth", dr["ForPeriodC9_YrMth"].ToString());                      
                        }
                        if (dr["ForPeriodC10_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriodC10_YrMth", dr["ForPeriodC10_YrMth"].ToString());                     
                        }
                        if (dr["ForPeriodC11_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriodC11_YrMth", dr["ForPeriodC11_YrMth"].ToString());                    
                        }
                        if (dr["ForPeriodC12_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriodC12_YrMth", dr["ForPeriodC12_YrMth"].ToString());                   
                        }
                        if (dr["ForPeriod1_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod1_YrMth", dr["ForPeriod1_YrMth"].ToString());                  
                        }
                        if (dr["ForPeriod2_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod2_YrMth", dr["ForPeriod2_YrMth"].ToString());                 
                        }
                        if (dr["ForPeriod3_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod3_YrMth", dr["ForPeriod3_YrMth"].ToString());                
                        }
                        if (dr["ForPeriod4_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod4_YrMth", dr["ForPeriod4_YrMth"].ToString());               
                        }
                        if (dr["ForPeriod5_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod5_YrMth", dr["ForPeriod5_YrMth"].ToString());              
                        }
                        if (dr["ForPeriod6_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod6_YrMth", dr["ForPeriod6_YrMth"].ToString());             
                        }
                        if (dr["ForPeriod7_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod7_YrMth", dr["ForPeriod7_YrMth"].ToString());            
                        }
                        if (dr["ForPeriod8_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod8_YrMth", dr["ForPeriod8_YrMth"].ToString());           
                        }
                        if (dr["ForPeriod9_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod9_YrMth", dr["ForPeriod9_YrMth"].ToString());          
                        }
                        if (dr["ForPeriod10_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod10_YrMth", dr["ForPeriod10_YrMth"].ToString());         
                        }
                        if (dr["ForPeriod11_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod11_YrMth", dr["ForPeriod11_YrMth"].ToString());        
                        }
                        if (dr["ForPeriod12_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod12_YrMth", dr["ForPeriod12_YrMth"].ToString());       
                        }
                        if (dr["ForPeriod13_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod13_YrMth", dr["ForPeriod13_YrMth"].ToString());      
                        }
                        if (dr["ForPeriod14_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod14_YrMth", dr["ForPeriod14_YrMth"].ToString());     
                        }
                        if (dr["ForPeriod15_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod15_YrMth", dr["ForPeriod15_YrMth"].ToString());    
                        }
                        if (dr["ForPeriod16_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod16_YrMth", dr["ForPeriod16_YrMth"].ToString());   
                        }
                        if (dr["ForPeriod17_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod17_YrMth", dr["ForPeriod17_YrMth"].ToString());  
                        }
                        if (dr["ForPeriod18_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod18_YrMth", dr["ForPeriod18_YrMth"].ToString()); 
                        }
                        if (dr["ForPeriod19_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod19_YrMth", dr["ForPeriod19_YrMth"].ToString());
                        }
                        if (dr["ForPeriod20_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod20_YrMth", dr["ForPeriod20_YrMth"].ToString());   
                        }
                        if (dr["ForPeriod21_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod21_YrMth", dr["ForPeriod21_YrMth"].ToString());  
                        }
                        if (dr["ForPeriod22_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod22_YrMth", dr["ForPeriod22_YrMth"].ToString()); 
                        }
                        if (dr["ForPeriod23_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod23_YrMth", dr["ForPeriod23_YrMth"].ToString());
                        }
                        if (dr["ForPeriod24_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod24_YrMth", dr["ForPeriod24_YrMth"].ToString());  
                        }
                        if (dr["ForPeriod_After_YrMth"] != System.DBNull.Value)
                        {
                            result.Add("ForPeriod_After_YrMth", dr["ForPeriod_After_YrMth"].ToString()); 
                        }
                        if (dr["Upload_Status"] != System.DBNull.Value)
                        {
                            result.Add("Upload_Status", dr["Upload_Status"].ToString());
                        }

                        if (dr["Total_Budget_Four_Mths_to_1"] != System.DBNull.Value)
                        {
                            result.Add("Total_Budget_Four_Mths_to_1", dr["Total_Budget_Four_Mths_to_1"].ToString());
                        }
                        if (dr["Cum_Four_Mths_to_1"] != System.DBNull.Value)
                        {
                            result.Add("Cum_Four_Mths_to_1", dr["Cum_Four_Mths_to_1"].ToString());
                        }
                        if (dr["Total_Twelve_Mths_to_1"] != System.DBNull.Value)
                        {
                            result.Add("Total_Twelve_Mths_to_1", dr["Total_Twelve_Mths_to_1"].ToString());
                        }

                        if (dr["Cum_Twelve_Mths_to_1"] != System.DBNull.Value)
                        {
                            result.Add("Cum_Twelve_Mths_to_1", dr["Cum_Twelve_Mths_to_1"].ToString());
                        }
                        if (dr["Total_Twelve_Mths_to_2"] != System.DBNull.Value)
                        {
                            result.Add("Total_Twelve_Mths_to_2", dr["Total_Twelve_Mths_to_2"].ToString());
                        }
                        if (dr["Total_All"] != System.DBNull.Value)
                        {
                            result.Add("Total_All", dr["Total_All"].ToString());
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
        private void Export_button_Click(object sender, EventArgs e)
        {
            #region export
            List<String> Project_Code_List = new List<String>();
            string date_time_string = "";
            string _PCMSconnectionString = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(_PCMSconnectionString))
            {
                try
                {
                    SqlCommand command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    //command.CommandText = "select DataField_Name, Col_No from dbo.Budget_CashFlow_Excel_ReportItem_RowCol where DataField_Name not in ('Budget_CashFlow_ReportItem_Name','Budget_CashFlow_ReportItem_Code','CumTo_ForPeriod','[START_ROW_NO]') and Budget_CashFlow_Type='" + cashflow_type + "'";
                   

                    DateTime dt = Convert.ToDateTime(Date_comboBox.SelectedValue.ToString());

                    if (dt.Month < 10)
                        date_time_string = dt.Year.ToString() + "0" + dt.Month.ToString();
                    else
                        date_time_string = dt.Year.ToString() + dt.Month.ToString();

                    command.CommandText = "select Budget_Project_No from dbo.Budget_CashFlow_Project where Budget_Project_No >= '" + ProjectCodeFrom_comboBox.SelectedValue.ToString() + "' and Budget_Project_No <= '" + ProjectCodeTo_comboBox.SelectedValue.ToString() + "' and YrEnd_CalPeriod = '" + date_time_string + "'";
                    connection.Open();
                    command.Connection = connection;
                    SqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        if (dr["Budget_Project_No"] != System.DBNull.Value)
                        {
                            Project_Code_List.Add(dr["Budget_Project_No"].ToString());
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

            object Missing = System.Type.Missing;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;
            Excel.Range oRng = null;
            Excel.Application oXL = new Excel.Application();
            Excel.Sheets xlSheets = null; 
            try
            {
                oXL.Visible = false;
                oXL.DisplayAlerts = false;

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                Boolean first = true;

                foreach (String project_code_item in Project_Code_List)
                {

                    if (first == true)
                    {
                        oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                        first = false;
                    }
                    else
                    {
                        xlSheets = oWB.Sheets as Excel.Sheets;
                        oSheet = (Excel._Worksheet)xlSheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                        oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    }
                    Dictionary<string, Dictionary<string, string[]>> Detail = ReadDetailRowFromDB(project_code_item, date_time_string);

                    //Dictionary<string, Dictionary<string, string[]>> DetailIncludingColTotal = ReadDetailIncludingColTotalRowFromDB(ProjectCode_comboBox.SelectedValue.ToString());


                    Dictionary<string, string> DetailCol = null;
                    Dictionary<string, string[]> DetailCol2 = null;
                    bool once = false;
                    // Display the ProgressBar control.
                    progressBar1.Visible = true;
                    // Set Minimum to 1 to represent the first file being copied.
                    progressBar1.Minimum = 1;
                    // Set Maximum to the total number of files to copy.
                    progressBar1.Maximum = Detail.Count+10;
                    // Set the initial value of the ProgressBar.
                    progressBar1.Value = 1;
                    // Set the Step property to a value of 1 to represent each file being copied.
                    progressBar1.Step = 1;

                    string YearEnd = "";
                    foreach (string Budget_CashFlow_ReportItem_Name in Detail.Keys)
                    {
                        Dictionary<string, string[]> data;
                        Detail.TryGetValue(Budget_CashFlow_ReportItem_Name, out data);

                        foreach (string column_name in data.Keys)
                        {
                            string[] location_and_value;
                            data.TryGetValue(column_name, out location_and_value);
                            string row = location_and_value[0];
                            string value = location_and_value[1];
                            string type = location_and_value[2];
                            string yrend_period = location_and_value[3];

                            if (once == false)
                            {
                                YearEnd = yrend_period;
                                DetailCol = ReadDetailColFromDB(type, yrend_period);
                                once = true;
                            }
                            /*
                            if (type == "H")
                            {
                                oRng = oSheet.get_Range(header_col + row, header_col + row);
                                oRng.set_Value(Missing, column_name);
                            }
                            */
                            string col = "";
                            if (DetailCol.Count > 0)
                                DetailCol.TryGetValue(column_name, out col);
                            if (col == "I")
                                Console.WriteLine("");
                            if (column_name == "ForPeriod21")
                                Console.WriteLine("");
                            if (!string.IsNullOrEmpty(col) && !string.IsNullOrEmpty(row))
                            {
                                oSheet.Cells[row, 1] = Budget_CashFlow_ReportItem_Name.Substring(0, Budget_CashFlow_ReportItem_Name.IndexOf("$"));

                                if (string.IsNullOrEmpty(value))
                                    oSheet.Cells[row, col] = System.String.Format("0", "##,###,###,##0.00");
                                else
                                    oSheet.Cells[row, col] = System.String.Format(value, "##,###,###,##0.00");
                                oRng = (Excel.Range)oSheet.Cells[row, col];
                                oRng.NumberFormatLocal = "#,##0.00_);(#,##0.00)";
                                oRng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                            }
                        }
                        progressBar1.PerformStep();
                    }
                    // Column Total added at 27 Jan 2012
                    once = false;
                    //Dictionary<string, string[]> DetailCol2 = null;
                    foreach (string Budget_CashFlow_ReportItem_Name in Detail.Keys)
                    {
                        Dictionary<string, string[]> data;
                        Detail.TryGetValue(Budget_CashFlow_ReportItem_Name, out data);

                        foreach (string column_name in data.Keys)
                        {
                            string[] location_and_value;
                            data.TryGetValue(column_name, out location_and_value);
                            string row = location_and_value[0];
                            string value = location_and_value[1];
                            string type = location_and_value[2];
                            string yrend_period = location_and_value[3];

                            if (once == false)
                            {
                                YearEnd = yrend_period;
                                DetailCol2 = ReadDetailColTotalFromDB(type, yrend_period);
                                once = true;
                            }
                            /*
                            if (type == "H")
                            {
                                oRng = oSheet.get_Range(header_col + row, header_col + row);
                                oRng.set_Value(Missing, column_name);
                            }
                            */

                            //if( DetailCol.Count > 0)
                            //DetailCol.TryGetValue(column_name, out col);
                            string col = "";
                            string typ = "";
                            string formula = "";
                            foreach (string k in DetailCol2.Keys)
                            {
                                string[] col_and_formula = new string[2];
                                DetailCol2.TryGetValue(k, out col_and_formula);
                                col = col_and_formula[0];
                                formula = col_and_formula[1];
                                typ = col_and_formula[2];

                                if (typ == "T")
                                {
                                    string[] temp = formula.Replace("[", "").Replace("]", "").Split('+');

                                    string col2 = "";
                                    string typ2 = "";
                                    string formula2 = "";
                                    string final_formula = "=SUM(";
                                    foreach (string k2 in DetailCol2.Keys)
                                    {
                                        string[] col_and_formula2 = new string[2];
                                        DetailCol2.TryGetValue(k2, out col_and_formula2);
                                        col2 = col_and_formula2[0];
                                        formula2 = col_and_formula2[1];
                                        typ2 = col_and_formula2[2];

                                        if (typ2 == "D" || typ2 == "T")
                                        {
                                            foreach (string id in temp)
                                            {
                                                if (k2 == id)
                                                {
                                                    if (final_formula.EndsWith("("))
                                                        final_formula = final_formula + col2 + row;
                                                    else
                                                        final_formula = final_formula + "," + col2 + row;
                                                }
                                            }
                                        }
                                    }
                                    final_formula = final_formula + ")";
                                    if (!string.IsNullOrEmpty(col) && !string.IsNullOrEmpty(row))
                                    {
                                        //oSheet.Cells[row, 1] = Budget_CashFlow_ReportItem_Name.Substring(0, Budget_CashFlow_ReportItem_Name.IndexOf("$"));
                                        oRng = oSheet.get_Range(col + row, col + row);
                                        oRng.Formula = final_formula;
                                    }
                                }
                            }
                        }
                    }

                    Dictionary<string, string> worksheet = ReadWorkSheetFromDB(ProjectCodeFrom_comboBox.SelectedValue.ToString());
                    string Budget_Project_On_Hand_Status = "";
                    worksheet.TryGetValue("Budget_Project_On_Hand_Status", out Budget_Project_On_Hand_Status);
                    string YrEnd_Date = "";
                    worksheet.TryGetValue("YrEnd_Date", out YrEnd_Date);
                    int year = Convert.ToDateTime(YrEnd_Date).Year;
                    int month = Convert.ToDateTime(YrEnd_Date).Month;
                    int day = Convert.ToDateTime(YrEnd_Date).Day;
                    string Year_End = "";
                    if (month < 10)
                        Year_End = year.ToString() + "0" + month.ToString();
                    else
                        Year_End = year.ToString() + month.ToString();


                    if (!string.IsNullOrEmpty(Budget_Project_On_Hand_Status) && !string.IsNullOrEmpty(YrEnd_Date))
                    {
                        //Header
                        Dictionary<string, string[]> HeaderRowCol = ReadHeaderRowColFromDB(Year_End, Budget_Project_On_Hand_Status);

                        foreach (string HeaderFieldName in worksheet.Keys)
                        {
                            string[] location = new string[3];
                            HeaderRowCol.TryGetValue(HeaderFieldName, out location);
                            if (location != null)
                            {
                                if (!string.IsNullOrEmpty(location[0]) && !string.IsNullOrEmpty(location[1]))
                                {
                                    //if (location[1] + location[0] == "E10")
                                        //Console.WriteLine("");
                                    oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);
                                    string value = "";
                                    worksheet.TryGetValue(HeaderFieldName, out value);
                                    DateTime date_result;
                                    if (DateTime.TryParse(value, out date_result) && location[2] != null)
                                    {
                                        if (HeaderFieldName == "CumTo_ForPeriod_YrMth" ||
                                            HeaderFieldName == "ForPeriodC9_YrMth" ||
                                            HeaderFieldName == "ForPeriodC10_YrMth" ||
                                            HeaderFieldName == "ForPeriodC11_YrMth" ||
                                            HeaderFieldName == "ForPeriodC12_YrMth" ||
                                            HeaderFieldName == "ForPeriod1_YrMth" ||
                                            HeaderFieldName == "ForPeriod2_YrMth" ||
                                            HeaderFieldName == "ForPeriod3_YrMth" ||
                                            HeaderFieldName == "ForPeriod4_YrMth" ||
                                            HeaderFieldName == "ForPeriod5_YrMth" ||
                                            HeaderFieldName == "ForPeriod6_YrMth" ||
                                            HeaderFieldName == "ForPeriod7_YrMth" ||
                                            HeaderFieldName == "ForPeriod8_YrMth" ||
                                            HeaderFieldName == "ForPeriod9_YrMth" ||
                                            HeaderFieldName == "ForPeriod10_YrMth" ||
                                            HeaderFieldName == "ForPeriod11_YrMth" ||
                                            HeaderFieldName == "ForPeriod12_YrMth" ||
                                            HeaderFieldName == "ForPeriod13_YrMth" ||
                                            HeaderFieldName == "ForPeriod14_YrMth" ||
                                            HeaderFieldName == "ForPeriod15_YrMth" ||
                                            HeaderFieldName == "ForPeriod16_YrMth" ||
                                            HeaderFieldName == "ForPeriod17_YrMth" ||
                                            HeaderFieldName == "ForPeriod18_YrMth" ||
                                            HeaderFieldName == "ForPeriod19_YrMth" ||
                                            HeaderFieldName == "ForPeriod20_YrMth" ||
                                            HeaderFieldName == "ForPeriod21_YrMth" ||
                                            HeaderFieldName == "ForPeriod22_YrMth" ||
                                            HeaderFieldName == "ForPeriod23_YrMth" ||
                                            HeaderFieldName == "ForPeriod24_YrMth" ||
                                            HeaderFieldName == "ForPeriod_After_YrMth" ||
                                            HeaderFieldName == "ForPeriod_After_YrMth" ||
                                            HeaderFieldName == "ForPeriod_After_YrMth"
                                            )
                                        {
                                            string y = date_result.ToString("yy");
                                            string m = date_result.ToString("MMM");

                                            oRng.set_Value(Missing, m + "-" + y);
                                        }
                                        else
                                        {
                                            if (location[2] == "D")
                                                oRng.set_Value(Missing, date_result.ToString("dd-MMM-yyyy"));
                                            else
                                            {
                                                oRng.set_Value(Missing, value.ToString());
                                                oRng.NumberFormatLocal = "#,##0.00_);(#,##0.00)";
                                                oRng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        oRng.set_Value(Missing, value);
                                        if (Convert.ToInt32(location[0]) < 10)
                                        {
                                            oRng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                            oRng.NumberFormatLocal = "#,##0.00_);(#,##0.00)";
                                        }
                                    }
                                }
                            }

                        }
                        // Header formula
                        Dictionary<string, string[]> HeaderRowColFormula = ReadHeaderRowColFormulaFromDB(Year_End, Budget_Project_On_Hand_Status);
                        int[] PeriodRow = ReadHeaderRowColPeriodRowFromDB(Year_End, Budget_Project_On_Hand_Status);

                        foreach (string id in HeaderRowColFormula.Keys)
                        {
                            string[] location = new string[3];
                            HeaderRowColFormula.TryGetValue(id, out location);
                            if (location != null)
                            {
                                if (!string.IsNullOrEmpty(location[0]) && !string.IsNullOrEmpty(location[1]))
                                {
                                    oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);
                                    Boolean found = false;
                                    for (int i = 0; i < PeriodRow.Count(); i++)
                                    {
                                        if (Convert.ToInt32(location[0]) == PeriodRow[i])
                                        {
                                            found = true;
                                        }
                                    }
                                    if (found == false)
                                    {
                                        oRng.set_Value(Missing, System.String.Format("0", "##,###,###,##0.00"));
                                        oRng.NumberFormatLocal = "#,##0.00_);(#,##0.00)";
                                    }
                                    string value = "";
                                    oRng.Formula = location[2];
                                }
                            }
                        }
                        /*
                        foreach (string HeaderFieldName in worksheet.Keys)
                        {
                            string[] location = new string[3];
                            HeaderRowCol.TryGetValue(HeaderFieldName, out location);
                            if (location != null)
                            {
                                if (!string.IsNullOrEmpty(location[0]) && !string.IsNullOrEmpty(location[1]))
                                {
                                    oRng = oSheet.get_Range(location[1] + location[0], location[1] + location[0]);
                                    string value = "";
                                    worksheet.TryGetValue(HeaderFieldName, out value);
                                    DateTime date_result;
                                    if (DateTime.TryParse(value, out date_result) && location[2] != null)
                                    {
                                        if (HeaderFieldName == "ForPeriodC9_YrMth"
                                            )
                                        {
                                            string y = date_result.ToString("yy");
                                            string m = date_result.ToString("MMM");

                                            oRng.set_Value(Missing, m + "-" + y);
                                        }
                                        else
                                            oRng.set_Value(Missing, date_result.ToString("dd-MMM-yyyy"));
                                    }
                                    else
                                        oRng.set_Value(Missing, value);
                                }
                            }
                        }*/
                    }
                    Dictionary<string, string[]> ReportItem_Detail_Header = ReadReportItem_Detail_Header(Budget_Project_On_Hand_Status);
                    foreach (string HeaderFieldName in ReportItem_Detail_Header.Keys)
                    {
                        string[] headername_location = new string[2];
                        ReportItem_Detail_Header.TryGetValue(HeaderFieldName, out headername_location);
                        if (headername_location != null)
                        {
                            if (!string.IsNullOrEmpty(headername_location[0]) && !string.IsNullOrEmpty(headername_location[1]))
                            {
                                oRng = oSheet.get_Range(headername_location[0] + headername_location[1], headername_location[0] + headername_location[1]);
                                string value = "";
                                oRng.set_Value(Missing, HeaderFieldName);
                            }
                        }
                    }
                    // Only deal with all cost except Total direct cost and Total indirect cost
                    Dictionary<string, string[]> ReportItem_Detail_Total = ReadReportItem_Detail_Total(Budget_Project_On_Hand_Status);
                    Dictionary<String, String> ForTotalConstructionCostCol = new Dictionary<String, String>();

                    List<Formula_Class> formulas = new List<Formula_Class>();
                    foreach (string HeaderFieldName in ReportItem_Detail_Total.Keys)
                    {
                        string[] headername_location = new string[4];
                        ReportItem_Detail_Total.TryGetValue(HeaderFieldName, out headername_location);
                        if (headername_location != null)
                        {
                            if (!string.IsNullOrEmpty(headername_location[0]) && !string.IsNullOrEmpty(headername_location[1]))
                            {
                                oRng = oSheet.get_Range(headername_location[0] + headername_location[1], headername_location[0] + headername_location[1]);
                                string value = "";
                                oRng.set_Value(Missing, HeaderFieldName);
                            }
                            if (headername_location.Count() > 2)
                            {
                                if (!string.IsNullOrEmpty(headername_location[2]))
                                {
                                    string[] total_cells = headername_location[2].Split('+');
                                    for (int i = 0; i < total_cells.Count(); i++)
                                    {
                                        total_cells[i] = total_cells[i].Replace("[", "").Replace("]", "");
                                    }
                                    string formula = "";
                                    once = false;
                                    foreach (string Budget_CashFlow_ReportItem_Name in Detail.Keys)
                                    {
                                        Dictionary<string, string[]> data;
                                        Detail.TryGetValue(Budget_CashFlow_ReportItem_Name, out data);

                                        foreach (string column_name in data.Keys)
                                        {
                                            string[] location_and_value;
                                            data.TryGetValue(column_name, out location_and_value);
                                            string row = location_and_value[0];
                                            string value = location_and_value[1];
                                            string type = location_and_value[2];
                                            string yrend_period = location_and_value[3];
                                            string seq = location_and_value[4];

                                            if (column_name == "ForPeriod21")
                                                Console.WriteLine("");
                                            if (headername_location[3] == type)
                                            {
                                                if (HeaderFieldName == "Total Cash Inflow")
                                                    Console.WriteLine("");
                                                if (once == false)
                                                {
                                                    YearEnd = yrend_period;
                                                    DetailCol = ReadDetailColFromDB(type, yrend_period);
                                                    once = true;
                                                }
                                                string col = "";
                                                if (DetailCol.Count > 0)
                                                    DetailCol.TryGetValue(column_name, out col);

                                                if (col == "F")
                                                    Console.WriteLine("");
                                                if (!string.IsNullOrEmpty(col))
                                                {
                                                    if (total_cells[0] == seq)
                                                    {
                                                        Formula_Class fc = new Formula_Class();

                                                        formula = "=Sum(" + col + row;
                                                        fc.formula = formula;
                                                        fc.col = col;
                                                        fc.row = headername_location[1];
                                                        fc.total_name = HeaderFieldName;


                                                        Boolean found = false;
                                                        foreach (Formula_Class f in formulas)
                                                        {
                                                            double double_result = 0;
                                                            if (Double.TryParse(f.formula.Substring(6, 1), out double_result) == true)
                                                            {
                                                                if (f.formula.Substring(5, 1) == col && f.total_name == HeaderFieldName)
                                                                    found = true;
                                                            }
                                                            else
                                                            {
                                                                if (f.formula.Substring(5, 2) == col && f.total_name == HeaderFieldName)
                                                                    found = true;
                                                            }
                                                        }
                                                        if (found == false)
                                                        {
                                                            formulas.Add(fc);
                                                            if (!ForTotalConstructionCostCol.ContainsValue(col))
                                                                ForTotalConstructionCostCol.Add(ForTotalConstructionCostCol.Count.ToString(), col);
                                                        }
                                                    }
                                                    if (total_cells[total_cells.Count() - 1] == seq)
                                                    {
                                                        foreach (Formula_Class fc in formulas)
                                                        {
                                                            double double_result = 0;
                                                            if (Double.TryParse(fc.formula.Substring(6, 1), out double_result) == true)
                                                            {
                                                                if (fc.formula.Substring(5, 1) == col && fc.total_name == HeaderFieldName)
                                                                    if (!fc.formula.EndsWith(")"))
                                                                        fc.formula = fc.formula + ":" + col + row + ")";
                                                            }
                                                            else
                                                            {
                                                                if (fc.formula.Substring(5, 2) == col && fc.total_name == HeaderFieldName)
                                                                    if (!fc.formula.EndsWith(")"))
                                                                        fc.formula = fc.formula + ":" + col + row + ")";
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    // Total's Total except total direct cost and indirect cost - extreme buster

                    foreach (string HeaderFieldName in ReportItem_Detail_Total.Keys)
                    {
                        string[] headername_location = new string[4];
                        ReportItem_Detail_Total.TryGetValue(HeaderFieldName, out headername_location);
                        if (headername_location != null)
                        {
                            if (!string.IsNullOrEmpty(headername_location[0]) && !string.IsNullOrEmpty(headername_location[1]))
                            {
                                oRng = oSheet.get_Range(headername_location[0] + headername_location[1], headername_location[0] + headername_location[1]);
                                string value = "";
                                oRng.set_Value(Missing, HeaderFieldName);
                            }
                            if (headername_location.Count() > 2)
                            {
                                if (!string.IsNullOrEmpty(headername_location[2]))
                                {
                                    string[] total_cells = headername_location[2].Replace("[", "").Replace("]", "").Split('+');

                                    string formula = "";
                                    once = false;
                                    DetailCol2 = ReadDetailColTotalTotalFromDB("BUDGET", Year_End);

                                    Dictionary<string, string> rowinfo = ReadTotalTotalRowInfoFromDB(Year_End, Budget_Project_On_Hand_Status);
                                    foreach (string k in DetailCol2.Keys)
                                    {
                                        string[] col = new string[3];
                                        DetailCol2.TryGetValue(k, out col);

                                        Formula_Class fc = new Formula_Class();

                                        formula = "=Sum(";
                                        for (int i = 0; i < total_cells.Count(); i++)
                                        {
                                            string row = "";
                                            rowinfo.TryGetValue(total_cells[i], out row);

                                            if (total_cells.Count() == 2)
                                            {
                                                if (formula.EndsWith("("))
                                                    formula = formula + col[0] + row;
                                                else
                                                    formula = formula + "," + col[0] + row;
                                            }
                                            else
                                            {
                                                if (formula.EndsWith("("))
                                                {
                                                    formula = formula + col[0] + row;
                                                    rowinfo.TryGetValue(total_cells[total_cells.Count() - 1], out row);
                                                    formula = formula + ":" + col[0] + row;
                                                }
                                            }
                                        }
                                        formula = formula + ")";
                                        fc.formula = formula;
                                        fc.col = col[0];
                                        fc.row = headername_location[1];
                                        fc.total_name = HeaderFieldName;

                                        Boolean found = false;
                                        foreach (Formula_Class f in formulas)
                                        {
                                            double double_result = 0;
                                            if (Double.TryParse(f.formula.Substring(6, 1), out double_result) == true)
                                            {
                                                if (f.formula.Substring(5, 1) == col[0] && f.total_name == HeaderFieldName)
                                                    found = true;
                                            }
                                            else
                                            {
                                                if (f.formula.Substring(5, 2) == col[0] && f.total_name == HeaderFieldName)
                                                    found = true;
                                            }
                                        }
                                        if (found == false)
                                        {
                                            try
                                            {
                                                formulas.Add(fc);
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine(ex.Message.ToString());
                                            }
                                            if (!ForTotalConstructionCostCol.ContainsValue(col[0]))
                                                ForTotalConstructionCostCol.Add(ForTotalConstructionCostCol.Count.ToString(), col[0]);
                                        }
                                    }
                                    Console.WriteLine("");
                                }
                            }
                        }
                    }
                    Console.WriteLine("");

                    // Only deal with Total direct cost and Total indirect cost
                    Dictionary<string, Dictionary<string, string[]>> TotalCost = ReadTotalCostFromDB(Year_End, Budget_Project_On_Hand_Status);
                    foreach (string HeaderFieldName in ReportItem_Detail_Total.Keys)
                    {
                        string[] headername_location = new string[4];
                        ReportItem_Detail_Total.TryGetValue(HeaderFieldName, out headername_location);
                        if (headername_location != null)
                        {
                            if (!string.IsNullOrEmpty(headername_location[0]) && !string.IsNullOrEmpty(headername_location[1]))
                            {
                                oRng = oSheet.get_Range(headername_location[0] + headername_location[1], headername_location[0] + headername_location[1]);
                                string value = "";
                                oRng.set_Value(Missing, HeaderFieldName);
                            }
                            if (headername_location.Count() > 2)
                            {
                                if (!string.IsNullOrEmpty(headername_location[2]))
                                {
                                    string[] total_cells = headername_location[2].Split('+');
                                    for (int i = 0; i < total_cells.Count(); i++)
                                    {
                                        total_cells[i] = total_cells[i].Replace("[", "").Replace("]", "");
                                    }
                                    string formula = "";
                                    once = false;
                                    foreach (string Budget_CashFlow_ReportItem_Name in TotalCost.Keys)
                                    {
                                        Dictionary<string, string[]> data;
                                        TotalCost.TryGetValue(Budget_CashFlow_ReportItem_Name, out data);

                                        foreach (string column_name in data.Keys)
                                        {
                                            string[] location_and_value;
                                            data.TryGetValue(column_name, out location_and_value);
                                            string row = location_and_value[0];
                                            string value = location_and_value[1];
                                            string type = location_and_value[2];
                                            string yrend_period = location_and_value[3];
                                            string seq = location_and_value[4];

                                            if (HeaderFieldName == "Total Construction Costs")
                                                Console.WriteLine("");
                                            if (headername_location[3] == type)
                                            {
                                                if (HeaderFieldName == "Total Construction Costs")
                                                    Console.WriteLine("");
                                                if (once == false)
                                                {
                                                    YearEnd = yrend_period;
                                                    DetailCol = ReadDetailColFromDB(type, yrend_period);
                                                    once = true;
                                                }
                                                string col = "";
                                                if (DetailCol.Count > 0)
                                                    DetailCol.TryGetValue(column_name, out col);

                                                if (col == "F")
                                                    Console.WriteLine("");
                                                if (!string.IsNullOrEmpty(col))
                                                {
                                                    if (total_cells[0] == seq)
                                                    {
                                                        string old_col = col;
                                                        foreach (string k in ForTotalConstructionCostCol.Keys)
                                                        {

                                                            Formula_Class fc = new Formula_Class();

                                                            ForTotalConstructionCostCol.TryGetValue(k, out col);
                                                            formula = "=Sum(" + col + row;
                                                            fc.formula = formula;
                                                            fc.col = col;
                                                            fc.row = headername_location[1];
                                                            fc.total_name = HeaderFieldName;


                                                            Boolean found = false;
                                                            foreach (Formula_Class f in formulas)
                                                            {
                                                                double double_result = 0;
                                                                if (Double.TryParse(f.formula.Substring(6, 1), out double_result) == true)
                                                                {
                                                                    if (f.formula.Substring(5, 1) == col && f.total_name == HeaderFieldName)
                                                                        found = true;
                                                                }
                                                                else
                                                                {
                                                                    if (f.formula.Substring(5, 2) == col && f.total_name == HeaderFieldName)
                                                                        found = true;
                                                                }
                                                            }
                                                            if (found == false)
                                                            {
                                                                try
                                                                {
                                                                    formulas.Add(fc);
                                                                }
                                                                catch (Exception ex)
                                                                {
                                                                    Console.WriteLine(ex.Message.ToString());
                                                                }
                                                            }
                                                        }
                                                        col = old_col;
                                                    }
                                                    if (total_cells[total_cells.Count() - 1] == seq)
                                                    {
                                                        foreach (string k in ForTotalConstructionCostCol.Keys)
                                                        {
                                                            ForTotalConstructionCostCol.TryGetValue(k, out col);
                                                            foreach (Formula_Class fc in formulas)
                                                            {
                                                                double double_result = 0;
                                                                if (Double.TryParse(fc.formula.Substring(6, 1), out double_result) == true)
                                                                {
                                                                    if (fc.formula.Substring(5, 1) == col && fc.total_name == HeaderFieldName)
                                                                        if (!fc.formula.EndsWith(")"))
                                                                            fc.formula = fc.formula + "," + col + row + ")"; // differece at ","
                                                                }
                                                                else
                                                                {
                                                                    if (fc.formula.Substring(5, 2) == col && fc.total_name == HeaderFieldName)
                                                                        if (!fc.formula.EndsWith(")"))
                                                                            fc.formula = fc.formula + "," + col + row + ")"; // differece at ","
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    Console.WriteLine("");
                    foreach (Formula_Class fc in formulas)
                    {
                        oRng = oSheet.get_Range(fc.col + fc.row, fc.col + fc.row);
                        oRng.set_Value(Missing, System.String.Format("0", "##,###,###,##0.00"));
                        oRng.NumberFormatLocal = "#,##0.00_);(#,##0.00)";
                        oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                        oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                        if (fc.formula.Contains(":") || fc.formula.Contains(","))
                        {
                            try
                            {
                                oRng.set_Value(Missing, fc.formula);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("");
                            }
                        }
                        else
                        {
                            try
                            {
                                oRng.set_Value(Missing, fc.formula + ":" + fc.formula.Replace("=Sum(", "") + ")");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("");
                            }
                        }
                    }

                    Dictionary<string, string[]> HeaderDetail = ReadHeaderDetailFromDB(Year_End, Budget_Project_On_Hand_Status);
                    foreach (string HeaderFieldName in HeaderDetail.Keys)
                    {
                        string[] headername_location = new string[3];
                        HeaderDetail.TryGetValue(HeaderFieldName, out headername_location);
                        if (headername_location != null)
                        {
                            if (!string.IsNullOrEmpty(headername_location[0]) && !string.IsNullOrEmpty(headername_location[1]))
                            {
                                oRng = oSheet.get_Range(headername_location[1] + headername_location[0], headername_location[1] + headername_location[0]);
                                string value = "";
                                oRng.set_Value(Missing, headername_location[2]);
                            }
                        }
                    }
                    oRng = oSheet.get_Range("D3", "D3");// Hard code
                    oRng.set_Value(Missing, Budget_Project_On_Hand_Status);
                    oRng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                    string Title_Year_End = "";
                    Title_Year_End = day.ToString() + " " + Convert.ToDateTime(YrEnd_Date).ToString("MMM") + " " + year.ToString();

                    oRng = oSheet.get_Range("A1", "A1"); // Hard code
                    oRng.set_Value(Missing, "Project Budget For The Year Ended " + Title_Year_End);
                    if (Detail.Count == 0)
                        MessageBox.Show("No data");

                    oRng = oSheet.get_Range("A1", "Z65536");
                    oRng.EntireColumn.AutoFit();

                    oSheet.Columns.HorizontalAlignment = 1;
                    oSheet.Columns.VerticalAlignment = -4160;
                    oSheet.Columns.Orientation = 0;
                    oSheet.Columns.AddIndent = false;
                    oSheet.Columns.IndentLevel = 0;
                    oSheet.Columns.ShrinkToFit = false;
                    oSheet.Columns.ReadingOrder = -5002;
                    oSheet.Columns.MergeCells = false;
                    oSheet.Columns.Font.Name = "Arial";
                    oSheet.Columns.Font.FontStyle = "Normal";
                    oSheet.Columns.Font.Size = 10;
                    oSheet.Columns.Font.Strikethrough = false;
                    oSheet.Columns.Font.Superscript = false;
                    oSheet.Columns.Font.Subscript = false;
                    oSheet.Columns.Font.OutlineFont = false;
                    oSheet.Columns.Font.Shadow = false;
                    oSheet.Columns.Font.Underline = -4142;
                    oSheet.Columns.Font.ColorIndex = -4105;
                    if (oSheet.Cells[3, 1] == "Project Name")
                    {
                        oRng = oSheet.get_Range("A1", "A65536");
                        oRng.ColumnWidth = 114;
                    }

                    oRng = oSheet.get_Range("A1", "A65536");
                    //oXL.Visible = true;
                    for (int i = 0; i < 10; i++)
                        progressBar1.PerformStep();
                }

                if (first == false)
                {
                    //string fname = "D:\\Data\\My Documents\\Visual Studio 2008\\WebSites\\Cheques\\Reports\\ChequeDataEnquiry.xls";
                    string fname = Application.StartupPath + "\\Excel_" + DateTime.Now.ToString().Replace(":", "_").Replace("/", "-") + ".xls";
                    if (System.IO.File.Exists(fname))
                    {
                        System.IO.File.Delete(fname);
                        oWB.SaveCopyAs(fname);
                        MessageBox.Show("Export Success");
                    }
                    else
                    {
                        oWB.SaveCopyAs(fname);
                        MessageBox.Show("Export Success");
                    }
                }
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

                //this.Close();
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
            #endregion export
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'financial_Rpt_TestDataSet2.datelist' table. You can move, or remove it, as needed.
            this.datelistTableAdapter.Fill(this.financial_Rpt_TestDataSet2.datelist);
            // TODO: This line of code loads data into the 'financial_Rpt_TestDataSet.Budget_CashFlow_Project' table. You can move, or remove it, as needed.
            this.budget_CashFlow_ProjectTableAdapter.Fill(this.financial_Rpt_TestDataSet.Budget_CashFlow_Project);

        }
    }
}
