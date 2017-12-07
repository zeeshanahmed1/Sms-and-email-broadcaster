using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using System.Data;

namespace BulkEmailSMS
{
    class ConnectionClass
    {
        public string connectionString = @"Data Source=ALPHA;Initial Catalog=BulkSmsEmail;User ID=sa;Password=admin@sql";
       
        SqlConnection con;
        // SqlDataAdapter adapt;
        SqlCommand cmd;

        public void ExecuteQry(string SQlQry)
        {
            con = new SqlConnection(connectionString);
            con.Open();
            SqlCommand cmd = new SqlCommand(SQlQry, con);
            SqlDataReader dr;
            dr = cmd.ExecuteReader();
            con.Close();
        }
        public DataTable GetData(string sqlQuery)
        {
            con = new SqlConnection(connectionString);
            DataTable dt = null;
            try
            {
                con.Open();
                cmd = new SqlCommand(sqlQuery, con);
                dt = new DataTable();
                dt.Load(cmd.ExecuteReader());
            }
            catch (Exception)
            {
                MessageBox.Show("Can not open connection ! ");
            }

            finally
            {
                cmd.Dispose();
                con.Close();
            }

            return dt;
        }

        public void InsertRec(string tblName, string colName, string values)
        {

                ExecuteQry(@"INSERT INTO " + tblName + " ( " + colName + " ) VALUES( " + values + ") ");

           
        }
        public void DeleteRec(string tblName, string colName, string values)
        {
            ExecuteQry(@"DELETE FROM " + tblName + " WHERE  " + colName + "  = '" + values + "'");

        }


        public void UpdatePersonTableRec(string value1, string value2, string value3, string value4, string value5, string value6)
        {

            con = new SqlConnection(connectionString);
            cmd = new SqlCommand("update Contacts set ContactName=@name,Address=@address, Phone=@mno, Email=@mail , CategoryName=@catName where ContactID=@id", con);
            con.Open();
            cmd.Parameters.AddWithValue("@name", value1);
            cmd.Parameters.AddWithValue("@address", value2);
            cmd.Parameters.AddWithValue("@mno", value3);
            cmd.Parameters.AddWithValue("@mail", value4);
            cmd.Parameters.AddWithValue("@catName", value6);
            cmd.Parameters.AddWithValue("@id", value5);
            cmd.ExecuteNonQuery();
            con.Close();
        }

        public void cBox(DataTable dt, ComboBox cb)
        {

            cb.Items.Clear();

            cb.DisplayMember = "CategoryName";
            cb.ValueMember = "CategoryName";

            cb.DataSource = dt;

        }


        public void cBoxEmail(DataTable dt, ComboBox cb)
        {

            cb.Items.Clear();

            cb.DisplayMember = "EmailID";
            cb.ValueMember = "Password";

            cb.DataSource = dt;

        }


    }
}
