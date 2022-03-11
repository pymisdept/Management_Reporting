using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Sql;
using System.Data.Common;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Net.NetworkInformation;

namespace ReadCashFlow
{
    /*
    internal static class NativeMethods
    {
        [DllImport("kernel32.dll")]
        internal static extern Boolean AllocConsole();
    } */
    class Program
    {
        static void Main(string[] args)
        {
            //
            if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
            {
                Boolean networkpass = true;
                Ping pingSender = new Ping ();
                PingOptions options = new PingOptions ();

                // Use the default Ttl value which is 128,
                // but change the fragmentation behavior.
                options.DontFragment = true;

                // Create a buffer of 32 bytes of data to be transmitted.
                string data = "aaaaaaaaaaaaaaaaaaaaaaa";
                byte[] buffer = Encoding.ASCII.GetBytes (data);
                int timeout = 120;
                String[] prop = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString.Split(';');
                String ip = "";
                foreach (String p in prop)
                {
                    if(p.Contains("Data Source="))
                        ip = p.Replace("Data Source=","");
                }
                PingReply reply = pingSender.Send(ip, timeout, buffer, options);
                if (reply.Status != IPStatus.Success)
                {
                    networkpass = false;
                }

                prop = ConfigurationManager.ConnectionStrings["PCMSConnectionString"].ConnectionString.Split(';');
                ip = "";
                foreach (String p in prop)
                {
                    if(p.Contains("Data Source="))
                        ip = p.Replace("Data Source=","");
                }
                PingReply reply2 = pingSender.Send(ip, timeout, buffer, options);
                if (reply2.Status != IPStatus.Success)
                {
                    networkpass = false;
                }

                if (networkpass == true)
                {
                    if (args.Length >= 4)
                    {
                        //Console.WriteLine("test");
                        //Console.ReadKey();
                        string[] param = new string[5];

                        //MessageBox.Show(args[args.Count() - 1].ToString());

                        param[0] = args[1].Trim(); // one or many
                        param[1] = args[2].Trim(); // username
                        param[2] = args[3].Trim().Replace("$", " "); // path

                        //Console.WriteLine("param[0]:" + param[0]);
                        //Console.WriteLine("param[1]:" + param[1]);
                        //Console.WriteLine("param[2]:" + param[2]);
                        if (args.Length > 4)
                            param[3] = args[4].Trim().Replace("$", " "); // file name

                        param[4] = args[0].Trim(); // YearMonth
                        //Console.WriteLine("param[3]:" + param[3]);
                        //Console.ReadKey();

                        /*
                        param[0] = "many";
                        //param[0] = "one";
                        param[1] = "Martin";
                        param[2] = Environment.CurrentDirectory;
                        param[3] = "Template-Project (HK & Macau)4.xls";
                        param[4] = "201303";
                        */

                        Form1 view = new Form1(param);
                        view.TopMost = true;
                        view.StartPosition = FormStartPosition.CenterScreen;
                        Form1.CheckForIllegalCrossThreadCalls = false;

                        Form dummy = new Form();
                        //Application.Run(view);
                        try
                        {
                            //Application.Run(view);
                            view.ShowDialog(dummy);
                        }
                        catch (Exception ex)
                        {
                            Console.Write(ex.Message.ToString());
                        }
                        //Console.ReadKey();   
                    }
                    else
                    {
                        
                        string[] param = new string[5];
  
                        param[0] = "one";
                        //param[0] = "one";
                        param[1] = "Martin";
                        param[2] = Environment.CurrentDirectory;
                        param[3] = "Template-Project (HK & Macau)4.xls";
                        param[4] = "201303";
                        
                        Form1 view = new Form1(param);
                        view.TopMost = true;
                        view.StartPosition = FormStartPosition.CenterScreen;
                        Form1.CheckForIllegalCrossThreadCalls = false;

                        Form dummy = new Form();
                        //Application.Run(view);
                        try
                        {
                            //Application.Run(view);
                            view.ShowDialog(dummy);
                        }
                        catch (Exception ex)
                        {
                            Console.Write(ex.Message.ToString());
                        }
                        //Console.ReadKey();   
                    }
                }
                //return result;
            }
            else
            {
                Console.WriteLine("Network is not available");
            }
        }
    }
}
