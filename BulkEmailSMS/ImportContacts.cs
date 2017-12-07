using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BulkEmailSMS
{
    class ImportContacts
    {

        OleDbConnection Econ;
        SqlConnection con;

        string constr, Query, sqlconn;  
        private void ExcelConn(string FilePath)
        {

            constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", FilePath);
            Econ = new OleDbConnection(constr);

        }
        private void connection()
        {
            sqlconn = ConfigurationManager.ConnectionStrings["SqlCom"].ConnectionString;
            con = new SqlConnection(sqlconn);

        }


        public void InsertExcelRecords(string FilePath)
        {
            ExcelConn(FilePath);
           
            Query = string.Format("Select [ContactName],[Phone],[Email],[Address],[CategoryName] FROM [{0}]", "Sheet1$");
            OleDbCommand Ecom = new OleDbCommand(Query, Econ);
            Econ.Open();

            DataSet ds = new DataSet();
            OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
            Econ.Close();
            oda.Fill(ds);
            DataTable Exceldt = ds.Tables[0];
            connection();
            //creating object of SqlBulkCopy    
            SqlBulkCopy objbulk = new SqlBulkCopy(con);
            //assigning Destination table name    
            objbulk.DestinationTableName = "Contacts";
            //Mapping Table column    
            objbulk.ColumnMappings.Add("ContactName", "ContactName");
            objbulk.ColumnMappings.Add("Phone", "Phone");
            objbulk.ColumnMappings.Add("Email", "Email");
            objbulk.ColumnMappings.Add("Address", "Address");
            objbulk.ColumnMappings.Add("CategoryName", "CategoryName");

          

            //inserting Datatable Records to DataBase    
            con.Open();
            objbulk.WriteToServer(Exceldt);
            con.Close();
           MessageBox.Show("Successfully Imported");
        }


     
    }
}
