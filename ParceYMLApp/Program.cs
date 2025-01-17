﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParceYmlApp
{
    static class Program
    {
        public static bool InsertToDB;
        public static string PathExcelFile { get; internal set; }
        public static string PathFolderBase { get; internal set; }
        public static string connectionStr { get; set; }
        public static string PathXmlFile { get; set; }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmMain());
        }
    }
}
