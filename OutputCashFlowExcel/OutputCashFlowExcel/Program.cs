using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Configuration;
using System.Text;

namespace OutputCashFlowExcel
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
            {
                Boolean networkpass = true;
                Ping pingSender = new Ping();
                PingOptions options = new PingOptions();

                // Use the default Ttl value which is 128,
                // but change the fragmentation behavior.
                options.DontFragment = true;

                // Create a buffer of 32 bytes of data to be transmitted.
                string data = "aaaaaaaaaaaaaaaaaaaaaaa";
                byte[] buffer = Encoding.ASCII.GetBytes(data);
                int timeout = 120;
                String[] prop = ConfigurationManager.ConnectionStrings["PYMDBConnectionString"].ConnectionString.Split(';');
                String ip = "";
                foreach (String p in prop)
                {
                    if (p.Contains("Data Source="))
                        ip = p.Replace("Data Source=", "");
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
                    if (p.Contains("Data Source="))
                        ip = p.Replace("Data Source=", "");
                }
                PingReply reply2 = pingSender.Send(ip, timeout, buffer, options);
                if (reply2.Status != IPStatus.Success)
                {
                    networkpass = false;
                }

                if (networkpass == true)
                {
                    if (args.Length >= 3)
                    {
                        //Console.WriteLine("test");
                        //Console.ReadKey();
                        string[] param = new string[4];

                        //MessageBox.Show(args[args.Count() - 1].ToString());

                        param[0] = args[0].Trim();// user name
                        param[1] = args[1].Trim();// year month
                        param[2] = args[2].Trim();// from code
                        param[3] = args[3].Trim();// to code

                        Console.WriteLine("param[0]:" + param[0]);
                        Console.WriteLine("param[1]:" + param[1]);
                        Console.WriteLine("param[2]:" + param[2]);
                        Console.WriteLine("param[3]:" + param[3]);

                        //Console.ReadKey();

                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);
                        Form1 form = new Form1(param);

                        form.TopMost = true;
                        form.StartPosition = FormStartPosition.CenterScreen;
                        //form.Show();
                        Form1.CheckForIllegalCrossThreadCalls = false;
                        Form dummy = new Form();
                        //dummy.Visible = false;
                        form.ShowDialog(dummy);
                        //Application.Run(form);
                    }
                    else
                    {
                        //debug use
                        string[] param = new string[4];
                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);
                        Form1 form = new Form1(param);

                        form.TopMost = true;
                        form.StartPosition = FormStartPosition.CenterScreen;
                        Form1.CheckForIllegalCrossThreadCalls = false;
                        Form dummy = new Form();
                        form.ShowDialog(dummy);
                        //Application.Run(form);
                    }
                }
            }
            else
            {
                Console.WriteLine("Network is not available");
                Console.ReadKey();
            }
        }
    }
}
