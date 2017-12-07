using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BulkEmailSMS
{
    public partial class form_smpt : Form
    {
        ConnectionClass cn = new ConnectionClass();
        public form_smpt()
        {
            InitializeComponent();
            
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {

                if (tbUserID.Text == "" && tbPassword.Text == "")
                {
                    MessageBox.Show("Please enter EmailID and Password", "Note");

                }
                else
                {
                    new MailAddress(tbUserID.Text.Trim());
                    cn.InsertRec("[tbl_SMTP_ID's]", "EmailID, Password", "'" + tbUserID.Text.Trim() + "'  , '" + tbPassword.Text.Trim() + "'");
                    MessageBox.Show("Successsfully Saved","Note");
                    tbPassword.Clear();
                    tbUserID.Clear();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void form_smpt_Load(object sender, EventArgs e)
        {
            btnSave.TabStop = false;
            btnSave.FlatStyle = FlatStyle.Flat;
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btnSave.FlatAppearance.MouseOverBackColor = Color.Transparent;

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
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            mainForm f3 = new mainForm();
            f3.Show();
            this.Hide();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
