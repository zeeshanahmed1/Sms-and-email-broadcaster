using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;
using iTextSharp;
using iTextSharp.text.factories;
using iTextSharp.text.html;
using iTextSharp.text.rtf;
using System.Data.SqlClient;
using System.Net.Mail;

namespace BulkEmailSMS
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        ConnectionClass cn = new ConnectionClass();
        ImportContacts imp = new ImportContacts();
        
        int ID = 0;
        string appPath;
        string pass = null;

       
        private void btnAdd_Click(object sender, EventArgs e)
        {

            String category;
            try
            {
                if (tbFirstName.Text.Trim() == "" && tbAddress.Text == "" && tbEmail.Text == "")
                {
                    MessageBox.Show("Please fill all the fields");
                }
                else
                {
                    new MailAddress(tbEmail.Text.Trim());
                    if (tbCat.Text == "")
                    {
                        category = cbCat.SelectedValue.ToString();
                    }
                    else
                    {
                        category = tbCat.Text.Trim();
                    }

                    cn.InsertRec("Contacts", "ContactName, Email, Phone, Address, CategoryName", "'" + tbFirstName.Text.Trim() + "'  , '" + tbEmail.Text.Trim() + "' , '" + mtbMobileNo.Text.Trim() + "' ,  '" + tbAddress.Text.Trim() + "'  ,  '" + category + "'");
                    MessageBox.Show("Data Successfully added");
                    loadContactsDgv();
                    clearForm();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            desingbtns();
            loadContactsDgv();

            cn.cBox(cn.GetData("SELECT Distinct CategoryName FROM [Contacts] ORDER BY CategoryName ASC"), cbCat);
           

            DataGridViewImageButtonDeleteColumn delColumn = new DataGridViewImageButtonDeleteColumn();
            delColumn.Name = "btnDgvDelete";
            delColumn.HeaderText = "";
            dgvContactView.Columns.Add(delColumn);

            DataGridViewImageButtonEditColumn editRow = new DataGridViewImageButtonEditColumn();
            editRow.Name = "btnDgvEdit";
            editRow.HeaderText = "";
            dgvContactView.Columns.Add(editRow);
            


            dgvContactView.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvContactView.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
        }

        public void loadContactsDgv()
        {
            dgvContactView.DataSource = cn.GetData(@"SELECT ContactID AS [ID], ContactName AS [ContactName], Email AS [Email], Phone AS [Mobile No.], Address AS [Address],CategoryName AS [Category]    FROM Contacts ORDER BY ContactName ASC");


        }
        private void btnUpdate_Click(object sender, EventArgs e)
        {
            string cat;
            try
            {
                if (tbCat.Text == "")
                {
                     cat = cbCat.SelectedValue.ToString();
                }
                else
                {
                    cat = tbCat.Text.Trim();
                }


                new MailAddress(tbEmail.Text.Trim());
                cn.UpdatePersonTableRec(tbFirstName.Text.Trim(), tbAddress.Text.Trim(), mtbMobileNo.Text.Trim(), tbEmail.Text.Trim(), ID.ToString(), cat);
                MessageBox.Show("Contact Updated Successfully", "Note");
                loadContactsDgvByCat(cbCat.SelectedValue.ToString().Trim());
                clearForm();

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void dgvContactView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                var senderGrid = (DataGridView)sender;
                if (e.ColumnIndex == senderGrid.Columns["btnDgvEdit"].Index && e.RowIndex >= 0)
                {
                    // dissble add back and close buttons
                    this.btnAdd.Enabled = false;
                    tbFirstName.Text = dgvContactView.Rows[e.RowIndex].Cells["ContactName"].Value.ToString();
                    mtbMobileNo.Text = dgvContactView.Rows[e.RowIndex].Cells["Mobile No."].Value.ToString();
                    tbAddress.Text = dgvContactView.Rows[e.RowIndex].Cells["Address"].Value.ToString();
                    tbEmail.Text = dgvContactView.Rows[e.RowIndex].Cells["Email"].Value.ToString();
                    tbCat.Text = dgvContactView.Rows[e.RowIndex].Cells["Category"].Value.ToString();
                    ID = Convert.ToInt32(dgvContactView.Rows[e.RowIndex].Cells["ID"].Value.ToString());
                   

                }
                else if (e.ColumnIndex == senderGrid.Columns["btnDgvDelete"].Index && e.RowIndex >= 0)
                {
                    cn.DeleteRec("Contacts", "ContactID", dgvContactView.CurrentRow.Cells["ID"].Value.ToString().Trim());
                    MessageBox.Show("Contact Deleted Successfully", "Note");
                    loadContactsDgvByCat(cbCat.SelectedValue.ToString().Trim());

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void clearForm()
        {
            mtbMobileNo.Clear();
            tbFirstName.Clear();
            tbEmail.Clear();
            tbAddress.Clear();
            tbCat.Clear();

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            mainForm f3 = new mainForm();
            f3.Show();
            this.Hide();
        }

        public void desingbtns()
        {
            btnBack.TabStop = false;
            btnBack.FlatStyle = FlatStyle.Flat;
            btnBack.FlatAppearance.BorderSize = 0;
            btnBack.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            btnBack.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;

            btnAdd.TabStop = false;
            btnAdd.FlatStyle = FlatStyle.Flat;
            btnAdd.FlatAppearance.BorderSize = 0;
            btnAdd.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            btnAdd.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;

            btnUpdate.TabStop = false;
            btnUpdate.FlatStyle = FlatStyle.Flat;
            btnUpdate.FlatAppearance.BorderSize = 0;
            btnUpdate.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            btnUpdate.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;

            btnClose.TabStop = false;
            btnClose.FlatStyle = FlatStyle.Flat;
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            btnClose.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;


            btnGeneratePdf.TabStop = false;
            btnGeneratePdf.FlatStyle = FlatStyle.Flat;
            btnGeneratePdf.FlatAppearance.BorderSize = 0;
            btnGeneratePdf.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            btnGeneratePdf.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;


            btnImport.TabStop = false;
            btnImport.FlatStyle = FlatStyle.Flat;
            btnImport.FlatAppearance.BorderSize = 0;
            btnImport.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            btnImport.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;


        }

        private void btnGeneratePdf_Click(object sender, EventArgs e)
        {

           

            //try
            //{

               
            //    string appFolderPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);

            //    string resourcesFolderPath = Path.Combine(
            //    Directory.GetParent(appFolderPath).Parent.Name, @"Resources\");

            //    Document doc = new Document(iTextSharp.text.PageSize.LETTER, 10, 10, 42, 35);
            //    PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(resourcesFolderPath + "/List.pdf", FileMode.Create));
            //    doc.Open();
            //    doc.AddAuthor("ZEESHAN AHMED");
            //    BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
            //    iTextSharp.text.Font times = new iTextSharp.text.Font(bfTimes, 20, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLUE);
            //    Paragraph paragraph = new Paragraph("BULK EMAIL SMS APPLICATION\n\nCONTACT LIST\n\n", times);
            //    paragraph.Alignment = Element.ALIGN_CENTER;



            //    doc.Add(paragraph);

            //    PdfPTable table = new PdfPTable(dgvContactView.Columns.Count -2);
            //    for (int j = 0; j < dgvContactView.Columns.Count -2; j++)
            //    {
            //        table.AddCell(new Phrase(dgvContactView.Columns[j].HeaderText));
            //    }
            //    table.HeaderRows = 1;

            //    for (int i = 0; i < dgvContactView.Rows.Count; i++)
            //    {
            //        for (int k = 0; k < dgvContactView.Columns.Count; k++)
            //        {
            //            if (dgvContactView[k, i].Value != null)
            //            {
            //                table.AddCell(new Phrase(dgvContactView[k, i].Value.ToString()));

            //            }
            //        }
            //    }
            //    doc.Add(table);
            //    doc.Close();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
            //finally
            //{
            //    MessageBox.Show("successfull");
            //}



            string sSelectedPath = null;
            SqlConnection cnn;
            string connectionString = null;
            string sql = null;
            string data = null;
            int i = 0;
            int j = 0;

            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            connectionString = "Data Source=ALPHA;Initial Catalog=BulkSmsEmail;User ID=sa;Password=admin@sql";
            cnn = new SqlConnection(connectionString);
            cnn.Open();
            sql = "SELECT * FROM Contacts";
            SqlDataAdapter dscmd = new SqlDataAdapter(sql, cnn);
            DataSet ds = new DataSet();
            dscmd.Fill(ds);

            for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
            {
                for (j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet.Cells[i + 1, j + 1] = data;
                }
            }


            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Custom Description";

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                 sSelectedPath = fbd.SelectedPath + "\\informations.xls";
            }



            xlWorkBook.SaveAs(sSelectedPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            MessageBox.Show("Export Completed");



        }

        private void dgvContactView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridViewRow row = dgvContactView.Rows[e.RowIndex];
            row.DefaultCellStyle.BackColor = System.Drawing.Color.Thistle;
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog opFile = new OpenFileDialog();
            opFile.Title = "Select a File";
            //opFile.Filter = "jpg files (*.jpg)|*.jpg|All files (*.*)|*.*";

            string appFolderPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);

            string resourcesFolderPath = Path.Combine(
            Directory.GetParent(appFolderPath).Parent.Name, @"Resources\");



            if (Directory.Exists(resourcesFolderPath) == false)
            {
                Directory.CreateDirectory(resourcesFolderPath);
            }
            if (opFile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string file = opFile.FileName;
                   //  string adddress = resourcesFolderPath;

                    imp.InsertExcelRecords(file);


                }
                catch (Exception exp)
                {
                    MessageBox.Show("Unable to open file >>> " + exp.Message);
                }
            }
            else
            {
                opFile.Dispose();
            }



          

            loadContactsDgv();
        }

        private void cbCat_SelectedIndexChanged(object sender, EventArgs e)
        {
            pass = cbCat.SelectedValue.ToString().Trim();

            loadContactsDgvByCat(pass);
        }

        public void loadContactsDgvByCat(string catName)
        {
            dgvContactView.DataSource = cn.GetData(@"SELECT ContactID AS [ID], ContactName AS [ContactName], Email AS [Email], Phone AS [Mobile No.], Address AS [Address],CategoryName AS [Category]    FROM Contacts where CategoryName = '"  + catName+  "' ORDER BY ContactName ASC");


        }

       

    }
}