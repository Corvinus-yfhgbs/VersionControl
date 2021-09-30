
namespace UserMaintenance
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.listUsers = new System.Windows.Forms.ListBox();
            this.txtFullName = new System.Windows.Forms.TextBox();
            this.lblFullName = new System.Windows.Forms.Label();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnFileWrite = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // listUsers
            // 
            this.listUsers.FormattingEnabled = true;
            this.listUsers.Location = new System.Drawing.Point(23, 22);
            this.listUsers.Name = "listUsers";
            this.listUsers.Size = new System.Drawing.Size(259, 329);
            this.listUsers.TabIndex = 0;
            // 
            // txtFullName
            // 
            this.txtFullName.Location = new System.Drawing.Point(377, 22);
            this.txtFullName.Name = "txtFullName";
            this.txtFullName.Size = new System.Drawing.Size(180, 20);
            this.txtFullName.TabIndex = 1;
            // 
            // lblFullName
            // 
            this.lblFullName.AutoSize = true;
            this.lblFullName.Location = new System.Drawing.Point(300, 22);
            this.lblFullName.Name = "lblFullName";
            this.lblFullName.Size = new System.Drawing.Size(35, 13);
            this.lblFullName.TabIndex = 3;
            this.lblFullName.Text = "label1";
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(303, 60);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(254, 33);
            this.btnAdd.TabIndex = 5;
            this.btnAdd.Text = "button1";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnFileWrite
            // 
            this.btnFileWrite.Location = new System.Drawing.Point(303, 99);
            this.btnFileWrite.Name = "btnFileWrite";
            this.btnFileWrite.Size = new System.Drawing.Size(254, 33);
            this.btnFileWrite.TabIndex = 6;
            this.btnFileWrite.Text = "button1";
            this.btnFileWrite.UseVisualStyleBackColor = true;
            this.btnFileWrite.Click += new System.EventHandler(this.btnFileWrite_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(569, 363);
            this.Controls.Add(this.btnFileWrite);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.lblFullName);
            this.Controls.Add(this.txtFullName);
            this.Controls.Add(this.listUsers);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listUsers;
        private System.Windows.Forms.TextBox txtFullName;
        private System.Windows.Forms.Label lblFullName;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnFileWrite;
    }
}

