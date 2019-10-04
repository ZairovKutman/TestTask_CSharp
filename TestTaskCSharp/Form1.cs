using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestTaskCSharp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //connect to mysql server

            string connectionString;
            MySqlConnection mySqlConnection;

            connectionString = "server=localhost;Database=testtaskcsharp;Uid=root;Pwd=root;";
            mySqlConnection = new MySqlConnection(connectionString);
            mySqlConnection.Open();

            MessageBox.Show("Connection open ! ");

            MySqlCommand sqlCommand;
            MySqlDataReader sqlDataReader;
            String sql, output = "Your Data: \n\n", id = "", name = "", birthDate = "", phoneNumber = "", address = "", socialNumber = "";

            sql = "select* from Clients where SocialNumber ='12345543211234'";
            sqlCommand = new MySqlCommand(sql, mySqlConnection);
            sqlDataReader = sqlCommand.ExecuteReader();

            // Read data from mysql server

            while (sqlDataReader.Read())
            {
                id += sqlDataReader.GetValue(0);
                name += sqlDataReader.GetValue(1);
                birthDate += sqlDataReader.GetValue(2);
                phoneNumber += sqlDataReader.GetValue(3);
                address += sqlDataReader.GetValue(4);
                socialNumber += sqlDataReader.GetValue(5);

                output += id + "-" + name + "-" + birthDate + "-" + phoneNumber + "-" + address + "-" + socialNumber;
            }

            MessageBox.Show(output);
            sqlDataReader.Close();
            sqlCommand.Dispose();
            mySqlConnection.Close();

            //  Write and save to excel file

            var openDirectory = System.IO.Directory.GetCurrentDirectory() + @"\template\example.xlsx";
            var saveDirectory = System.IO.Directory.GetCurrentDirectory() + @"\result\example.xlsx";
            Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook sheet = excel.Workbooks.Open(openDirectory);
            Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;

            x.Range["B3"].Value = id;
            x.Range["B4"].Value = name;
            x.Range["B5"].Value = birthDate;
            x.Range["B6"].Value = phoneNumber;
            x.Range["B7"].Value = address;
            x.Range["B8"].Value = socialNumber;

            x.Range["D4"].Value = id;
            x.Range["E4"].Value = name;
            x.Range["F4"].Value = birthDate;
            x.Range["G4"].Value = phoneNumber;
            x.Range["H4"].Value = address;
            x.Range["I4"].Value = socialNumber + " ";

            MessageBox.Show("Your data successful wrote to folder result/example.xlsx ");

            sheet.SaveAs(saveDirectory);
            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }
    }
}
