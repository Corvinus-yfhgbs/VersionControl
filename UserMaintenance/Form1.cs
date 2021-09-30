using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using UserMaintenance.Entities;

namespace UserMaintenance
{
    public partial class Form1 : Form
    {
        BindingList<User> users = new BindingList<User>();
        public Form1()
        {
            InitializeComponent();
            lblFullName.Text = Resource1.FullName; // label1
            btnAdd.Text = Resource1.Add; // button1
            btnFileWrite.Text = Resource1.FileWrite;

            listUsers.DataSource = users;
            listUsers.ValueMember = "ID";
            listUsers.DisplayMember = "FullName";
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            var u = new User()
            {
                FullName = txtFullName.Text,
            };
            users.Add(u);
        }

        private void btnFileWrite_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.FileName = "NevLista.txt";
            saveFile.InitialDirectory = Application.StartupPath;
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                StreamWriter streamWriter = new StreamWriter(saveFile.OpenFile());
                for (int i = 0; i < listUsers.Items.Count; i++)
                {
                    streamWriter.WriteLine(users[i].ID + " - " + users[i].FullName);
                }

                streamWriter.Dispose();
                streamWriter.Close();
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            users.Remove((User)listUsers.SelectedItem);
        }
    }
}
