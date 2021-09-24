using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Nsoft
{
    class MyData
    {
        public static OleDbDataAdapter oledbadapter1;
        public static OleDbConnection oledbconnection1;

        public static DataTable dtmainKassa;
        public static DataTable dtmainSifarisler;
        public static DataTable dtmainQrafik;
        public static DataTable dtmainKredit;
        public static DataTable dtmainParol;
        public static DataTable dtmainLisenziya;

        public static void CreateConnection(string bazaName)
        {
            oledbconnection1 = new OleDbConnection();
            oledbadapter1 = new OleDbDataAdapter();
            oledbconnection1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + bazaName + "'";
        }    //elaqe yaratmaq

        public static void selectCommand(String bazaName, String commandText)
        {
            CreateConnection(bazaName);
            oledbadapter1.SelectCommand = new OleDbCommand();
            oledbadapter1.SelectCommand.Connection = oledbconnection1;
            oledbadapter1.SelectCommand.CommandText = commandText;
        }

        public static void insertCommand(String bazaName, String commandText)
        {
            CreateConnection(bazaName);
            oledbadapter1.InsertCommand = new OleDbCommand();
            oledbadapter1.InsertCommand.Connection = oledbconnection1;
            oledbconnection1.Open();
            oledbadapter1.InsertCommand.CommandText = commandText;
            oledbadapter1.InsertCommand.ExecuteNonQuery();
            oledbconnection1.Close();
        }
        public static void deleteCommand(String bazaName, String commandText)
        {
            CreateConnection(bazaName);
            oledbadapter1.DeleteCommand = new OleDbCommand();
            oledbadapter1.DeleteCommand.Connection = oledbconnection1;
            oledbconnection1.Open();
            oledbadapter1.DeleteCommand.CommandText = commandText;
            oledbadapter1.DeleteCommand.ExecuteNonQuery();
            oledbconnection1.Close();
        }

        public static void updateCommand(String bazaName, String commandText)
        {
            CreateConnection(bazaName);
            oledbadapter1.UpdateCommand = new OleDbCommand();
            oledbadapter1.UpdateCommand.Connection = oledbconnection1;
            oledbconnection1.Open();
            oledbadapter1.UpdateCommand.CommandText = commandText;
            oledbadapter1.UpdateCommand.ExecuteNonQuery();
            oledbconnection1.Close();
        }

        public static string appInfo()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            string result = "Company Name: " + fvi.CompanyName
                + Environment.NewLine + "Product Name: " + fvi.ProductName
                + Environment.NewLine + "File Location : " + fvi.FileName
                + Environment.NewLine + "Product Version: " + fvi.ProductVersion
                + Environment.NewLine
                + Environment.NewLine + "Comments: " + fvi.Comments
                + Environment.NewLine
                + Environment.NewLine + fvi.LegalCopyright;

            return result;
        }

    }
}
