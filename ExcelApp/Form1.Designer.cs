namespace ExcelApp
{
    partial class ExcelApp
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExcelApp));
            this.btnLoad = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.btnRead = new System.Windows.Forms.Button();
            this.btnWrite = new System.Windows.Forms.Button();
            this.txtFirstCell = new System.Windows.Forms.TextBox();
            this.labelFirstCell = new System.Windows.Forms.Label();
            this.btnCreate = new System.Windows.Forms.Button();
            this.btnWriteCells = new System.Windows.Forms.Button();
            this.btnProtect = new System.Windows.Forms.Button();
            this.buttonDeleteFile = new System.Windows.Forms.Button();
            this.btnDestination = new System.Windows.Forms.Button();
            this.textBoxFilename = new System.Windows.Forms.TextBox();
            this.labelFilename = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxDestination = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnLoad
            // 
            this.btnLoad.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLoad.Location = new System.Drawing.Point(88, 143);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(130, 41);
            this.btnLoad.TabIndex = 0;
            this.btnLoad.Text = "Load Excel File";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtFileName);
            this.groupBox1.Location = new System.Drawing.Point(32, 28);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(400, 105);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Selection of the Excel file";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Select the Excel file";
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(28, 50);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(351, 20);
            this.txtFileName.TabIndex = 1;
            // 
            // btnRead
            // 
            this.btnRead.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRead.Location = new System.Drawing.Point(233, 143);
            this.btnRead.Name = "btnRead";
            this.btnRead.Size = new System.Drawing.Size(143, 41);
            this.btnRead.TabIndex = 2;
            this.btnRead.Text = "Read the first cell";
            this.btnRead.UseVisualStyleBackColor = true;
            this.btnRead.Click += new System.EventHandler(this.btnRead_Click);
            // 
            // btnWrite
            // 
            this.btnWrite.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnWrite.Location = new System.Drawing.Point(51, 319);
            this.btnWrite.Name = "btnWrite";
            this.btnWrite.Size = new System.Drawing.Size(207, 49);
            this.btnWrite.TabIndex = 3;
            this.btnWrite.Text = "Write in the first cell";
            this.btnWrite.UseVisualStyleBackColor = true;
            this.btnWrite.Click += new System.EventHandler(this.btnWrite_Click);
            // 
            // txtFirstCell
            // 
            this.txtFirstCell.Location = new System.Drawing.Point(124, 42);
            this.txtFirstCell.Name = "txtFirstCell";
            this.txtFirstCell.Size = new System.Drawing.Size(100, 20);
            this.txtFirstCell.TabIndex = 4;
            // 
            // labelFirstCell
            // 
            this.labelFirstCell.AutoSize = true;
            this.labelFirstCell.Location = new System.Drawing.Point(23, 42);
            this.labelFirstCell.Name = "labelFirstCell";
            this.labelFirstCell.Size = new System.Drawing.Size(95, 13);
            this.labelFirstCell.TabIndex = 5;
            this.labelFirstCell.Text = "Text in the first cell";
            // 
            // btnCreate
            // 
            this.btnCreate.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreate.Location = new System.Drawing.Point(583, 164);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(75, 41);
            this.btnCreate.TabIndex = 6;
            this.btnCreate.Text = "Create";
            this.btnCreate.UseVisualStyleBackColor = true;
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // btnWriteCells
            // 
            this.btnWriteCells.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnWriteCells.Location = new System.Drawing.Point(177, 393);
            this.btnWriteCells.Name = "btnWriteCells";
            this.btnWriteCells.Size = new System.Drawing.Size(139, 49);
            this.btnWriteCells.TabIndex = 6;
            this.btnWriteCells.Text = "Write Multiple Cells";
            this.btnWriteCells.UseVisualStyleBackColor = true;
            this.btnWriteCells.Click += new System.EventHandler(this.btnWriteCells_Click);
            // 
            // btnProtect
            // 
            this.btnProtect.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProtect.Location = new System.Drawing.Point(458, 302);
            this.btnProtect.Name = "btnProtect";
            this.btnProtect.Size = new System.Drawing.Size(139, 50);
            this.btnProtect.TabIndex = 6;
            this.btnProtect.Text = "Protect Excel file";
            this.btnProtect.UseVisualStyleBackColor = true;
            this.btnProtect.Click += new System.EventHandler(this.btnProtect_Click);
            // 
            // buttonDeleteFile
            // 
            this.buttonDeleteFile.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonDeleteFile.Location = new System.Drawing.Point(32, 393);
            this.buttonDeleteFile.Name = "buttonDeleteFile";
            this.buttonDeleteFile.Size = new System.Drawing.Size(139, 49);
            this.buttonDeleteFile.TabIndex = 6;
            this.buttonDeleteFile.Text = "Delete the first worksheet";
            this.buttonDeleteFile.UseVisualStyleBackColor = true;
            this.buttonDeleteFile.Click += new System.EventHandler(this.buttonDeleteFile_Click);
            // 
            // btnDestination
            // 
            this.btnDestination.Image = ((System.Drawing.Image)(resources.GetObject("btnDestination.Image")));
            this.btnDestination.Location = new System.Drawing.Point(275, 79);
            this.btnDestination.Name = "btnDestination";
            this.btnDestination.Size = new System.Drawing.Size(40, 29);
            this.btnDestination.TabIndex = 6;
            this.btnDestination.UseVisualStyleBackColor = true;
            this.btnDestination.Click += new System.EventHandler(this.btnDestination_Click);
            // 
            // textBoxFilename
            // 
            this.textBoxFilename.Location = new System.Drawing.Point(156, 43);
            this.textBoxFilename.Name = "textBoxFilename";
            this.textBoxFilename.Size = new System.Drawing.Size(100, 20);
            this.textBoxFilename.TabIndex = 7;
            // 
            // labelFilename
            // 
            this.labelFilename.AutoSize = true;
            this.labelFilename.Location = new System.Drawing.Point(44, 43);
            this.labelFilename.Name = "labelFilename";
            this.labelFilename.Size = new System.Drawing.Size(49, 13);
            this.labelFilename.TabIndex = 8;
            this.labelFilename.Text = "Filename";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(41, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Destination of the file";
            // 
            // textBoxDestination
            // 
            this.textBoxDestination.Location = new System.Drawing.Point(156, 79);
            this.textBoxDestination.Name = "textBoxDestination";
            this.textBoxDestination.Size = new System.Drawing.Size(100, 20);
            this.textBoxDestination.TabIndex = 10;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.labelFilename);
            this.groupBox2.Controls.Add(this.textBoxDestination);
            this.groupBox2.Controls.Add(this.btnDestination);
            this.groupBox2.Controls.Add(this.textBoxFilename);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(450, 28);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(343, 130);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Create Excel File";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.txtFirstCell);
            this.groupBox3.Controls.Add(this.labelFirstCell);
            this.groupBox3.Location = new System.Drawing.Point(32, 221);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(250, 92);
            this.groupBox3.TabIndex = 12;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Write in the first cell";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 26.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(527, 391);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(235, 42);
            this.label3.TabIndex = 13;
            this.label3.Text = "ExcelAppSharp";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(458, 378);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(64, 64);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 14;
            this.pictureBox1.TabStop = false;
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Location = new System.Drawing.Point(83, 29);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.PasswordChar = '*';
            this.textBoxPassword.Size = new System.Drawing.Size(100, 20);
            this.textBoxPassword.TabIndex = 15;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(24, 29);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "Password";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.textBoxPassword);
            this.groupBox4.Location = new System.Drawing.Point(458, 221);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(200, 75);
            this.groupBox4.TabIndex = 17;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Protect the Excel file";
            // 
            // ExcelApp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(807, 459);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.buttonDeleteFile);
            this.Controls.Add(this.btnProtect);
            this.Controls.Add(this.btnWriteCells);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.btnWrite);
            this.Controls.Add(this.btnRead);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "ExcelApp";
            this.Text = "ExcelAppSharp";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Button btnRead;
        private System.Windows.Forms.Button btnWrite;
        private System.Windows.Forms.TextBox txtFirstCell;
        private System.Windows.Forms.Label labelFirstCell;
        private System.Windows.Forms.Button btnCreate;
        private System.Windows.Forms.Button btnWriteCells;
        private System.Windows.Forms.Button btnProtect;
        private System.Windows.Forms.Button buttonDeleteFile;
        private System.Windows.Forms.Button btnDestination;
        private System.Windows.Forms.TextBox textBoxFilename;
        private System.Windows.Forms.Label labelFilename;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxDestination;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox textBoxPassword;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox4;
    }
}

