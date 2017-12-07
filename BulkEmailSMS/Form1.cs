using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Net;
using System.IO;



namespace BulkEmailSMS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        ConnectionClass cn = new ConnectionClass();
        int b = 0;
        string pass;
        String cat;
        string smtpPortName = "";
        string smtpPortNo  = "";
        string iName;
        string filepath;
        private void Form1_Load(object sender, EventArgs e)
        {

            btnSendEmail.Enabled = false;
            cn.cBoxEmail(cn.GetData("SELECT EmailID,Password FROM [tbl_SMTP_ID's] ORDER BY EmailID ASC"), cmbEmailID);

            designbtns();
            LoadGrid();
            cn.cBox(cn.GetData("SELECT Distinct CategoryName FROM [Contacts] ORDER BY CategoryName ASC"), cbCat);
           

        }

        public void LoadGrid()
        {

            dgvSendContactView.DataSource = cn.GetData(@"SELECT ContactName AS [ContactName], Email AS [Email], Phone AS [Mobile No.], Address AS [Address],CategoryName AS [Category]  FROM Contacts ORDER BY ContactName ASC");

        }


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                
                if(richTextBox1.Text == "" && textBox2.Text == "") 
                {
                    MessageBox.Show("Please Enter Subject and Body","Error");
                
                }
                else
                {

                    string selected = string.Empty;
                    String selectedName = string.Empty;
                    foreach (DataGridViewRow row in dgvSendContactView.Rows)
                    {

                        bool isSelected = Convert.ToBoolean(row.Cells["chk"].Value);
                        if (isSelected)
                        {
                            b += 1;
                            selected += Environment.NewLine;
                            selected += row.Cells["Email"].Value.ToString();

                            selectedName += Environment.NewLine;
                            selectedName += row.Cells["ContactName"].Value.ToString();

                        }
                        
                    }
                    if (selected == "")
                    {
                        MessageBox.Show("Please Select Contacts to Send Email");
                    }



                    pbrSendEmail.Minimum = 0;
                    pbrSendEmail.Maximum = b;

                    //  MessageBox.Show("Selected Values" + message + b);
                    char[] delimiters = new char[] { '\r', '\n' };
                    string[] words = selected.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    string[] names = selectedName.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < b; i++)
                    {
                        #region Send Mail

                        int c = i + 1;
                        MailMessage mail = new MailMessage();
                        SmtpClient SmtpServer = new SmtpClient("" + smtpPortName.ToString().Trim() + "");
                        mail.From = new MailAddress("" + cmbEmailID.Text.Trim() + "");
                        mail.To.Add(words[i]);
                        //>>>> for CC                        mail.CC.Add(textBox2.Text);
                        

                        if (tbAttachementFilePath.Text == "")
                        {


                        }
                        else
                        {
                            mail.Attachments.Add(new Attachment("" + filepath.ToString().Trim() + ""));

                        }
                        
                        
                        
                        pbrSendEmail.Value = c;

                        mail.Subject = textBox2.Text.Trim();

                        string dear = "Dear " + names[i] + ",\n";

                        mail.Body = dear + richTextBox1.Text.Trim();
                        SmtpServer.Port = int.Parse(smtpPortNo.ToString().Trim());
                        SmtpServer.Credentials = new System.Net.NetworkCredential("" + cmbEmailID.Text.Trim() + "", "" + pass.ToString() + "");
                        SmtpServer.EnableSsl = true;
                        SmtpServer.Send(mail);

                        if (i == b - 1)
                        {
                            //MessageBox.Show("Mail Successfully Sent");
                            DialogResult result = MessageBox.Show("Mail Successfully Sent", "Congatulations!", MessageBoxButtons.OK);
                            if (result == DialogResult.OK)
                            {
                                textBox2.Clear();
                                richTextBox1.Clear();
                                tbAttachementFilePath.Clear();
                                pbrSendEmail.Value = 0;
                                break;
                            }
                            else if (result == DialogResult.No)
                            {
                                //no...
                            }
                        }
                    }
                        #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        
        }

        private void dgvContactView_SelectionChanged(object sender, EventArgs e)
        {
            
        }
        private void button2_Click(object sender, EventArgs e)
        {
           
           
            
            
        }

        private void dgvContactView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
           
        }

        private void dgvSendContactView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        public void designbtns()
        {
          
            btnSendEmail.TabStop = false;
            btnSendEmail.FlatStyle = FlatStyle.Flat;
            btnSendEmail.FlatAppearance.BorderSize = 0;
            btnSendEmail.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btnSendEmail.FlatAppearance.MouseOverBackColor = Color.Transparent;

            btnBack.TabStop = false;
            btnBack.FlatStyle = FlatStyle.Flat;
            btnBack.FlatAppearance.BorderSize = 0;
            btnBack.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btnBack.FlatAppearance.MouseOverBackColor = Color.Transparent;


            btnClose.TabStop = false;
            btnClose.FlatStyle = FlatStyle.Flat;
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btnClose.FlatAppearance.MouseOverBackColor = Color.Transparent;


            btnConfigureSmtp.TabStop = false;
            btnConfigureSmtp.FlatStyle = FlatStyle.Flat;
            btnConfigureSmtp.FlatAppearance.BorderSize = 0;
            btnConfigureSmtp.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btnConfigureSmtp.FlatAppearance.MouseOverBackColor = Color.Transparent;

            btnBrowseFile.TabStop = false;
            btnBrowseFile.FlatStyle = FlatStyle.Flat;
            btnBrowseFile.FlatAppearance.BorderSize = 0;
            btnBrowseFile.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btnBrowseFile.FlatAppearance.MouseOverBackColor = Color.Transparent;
        
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

        private void cmbEmailID_SelectedIndexChanged(object sender, EventArgs e)
        {
          
        }

        private void dgvSendContactView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridViewRow row = dgvSendContactView.Rows[e.RowIndex];
            row.DefaultCellStyle.BackColor = Color.Thistle;
        }

        private void pbrSendEmail_Click(object sender, EventArgs e)
        {

        }

        private void dgvSendContactView_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgvSendContactView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            foreach (DataGridViewRow dgv in dgvSendContactView.Rows)
            {
                dgv.Cells["chk"].Value = true;
            }
        }

        private void cmbEmailID_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            pass = cmbEmailID.SelectedValue.ToString().Trim();
        }

        private void btnConfigureSmtp_Click(object sender, EventArgs e)
        {
            if (tbSmtpServer.Text == "" && mtbSmtpPortNo.Text == "")
            {
                MessageBox.Show("Please enter Smtp Server and Port No.");
               
                
            }
            else
            {
                 smtpPortName = tbSmtpServer.Text.Trim();
                 smtpPortNo = mtbSmtpPortNo.Text.Trim();
                 btnSendEmail.Enabled = true;
            }
        }

        private void btnBrowseFile_Click(object sender, EventArgs e)
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


                    tbAttachementFilePath.Text = appFolderPath.ToString().Trim();
                    string ext = Path.GetExtension(opFile.FileName);
                    iName = cmbEmailID.Text.Trim() + "_" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + "_" + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + ext;
                    filepath = opFile.FileName;

                    File.Copy(filepath, resourcesFolderPath + iName);
                  
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Unable to open file " + exp.Message);
                }
            }
            else
            {
                opFile.Dispose();
            }
        }

        private void dgvSendContactView_CellFormatting_1(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridViewRow row = dgvSendContactView.Rows[e.RowIndex];
            row.DefaultCellStyle.BackColor = System.Drawing.Color.Thistle;
        }

        private void cbCat_SelectedIndexChanged(object sender, EventArgs e)
        {
            cat = cbCat.SelectedValue.ToString().Trim();

            loadContactsDgvByCat(cat);
        }


        public void loadContactsDgvByCat(string catName)
        {

            dgvSendContactView.DataSource = cn.GetData(@"SELECT ContactID AS [ID], ContactName AS [ContactName], Email AS [Email], Phone AS [Mobile No.], Address AS [Address],CategoryName AS [Category]    FROM Contacts where CategoryName = '" + catName + "' ORDER BY ContactName ASC");

        }

       
    }
}