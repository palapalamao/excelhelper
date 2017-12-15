using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace OverTimeStatistics
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            DateTime dt = DateTime.Now;
            if (dt.Year > 2017 && dt.Month >=3)
            {
                MessageBox.Show("异常退出，联系作者1160744812@qq.com");
                Application.Exit();
                System.Environment.Exit(0);
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Excelsplit());
        }



    }


     


}
