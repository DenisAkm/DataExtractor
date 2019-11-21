namespace TableDataExtractor
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxValue = new System.Windows.Forms.TextBox();
            this.listBoxSource = new System.Windows.Forms.ListBox();
            this.contextMenuStripSource = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.addFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonUp = new System.Windows.Forms.Button();
            this.buttonDown = new System.Windows.Forms.Button();
            this.buttonUpVal = new System.Windows.Forms.Button();
            this.buttonDownVal = new System.Windows.Forms.Button();
            this.listBoxValue = new System.Windows.Forms.ListBox();
            this.contextMenuStripValue = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.clearToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBoxDelimeter = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.comboBoxStile = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.numericUpDownColVal = new System.Windows.Forms.NumericUpDown();
            this.label7 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.textBoxDirectory = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.textBoxTableName = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.numericUpDownDecPlaces = new System.Windows.Forms.NumericUpDown();
            this.comboBoxSeparator = new System.Windows.Forms.ComboBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.tableLayoutPanel1.SuspendLayout();
            this.contextMenuStripSource.SuspendLayout();
            this.contextMenuStripValue.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownColVal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownDecPlaces)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.button1, 2);
            this.button1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(508, 292);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(265, 40);
            this.button1.TabIndex = 0;
            this.button1.Text = "Extract";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 7;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.7586F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 27.0388F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 8F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 19.89383F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 1.131691F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.17708F));
            this.tableLayoutPanel1.Controls.Add(this.label2, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this.textBoxValue, 3, 10);
            this.tableLayoutPanel1.Controls.Add(this.listBoxSource, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.buttonUp, 0, 12);
            this.tableLayoutPanel1.Controls.Add(this.buttonDown, 1, 12);
            this.tableLayoutPanel1.Controls.Add(this.button1, 5, 12);
            this.tableLayoutPanel1.Controls.Add(this.buttonUpVal, 3, 11);
            this.tableLayoutPanel1.Controls.Add(this.buttonDownVal, 3, 12);
            this.tableLayoutPanel1.Controls.Add(this.listBoxValue, 3, 1);
            this.tableLayoutPanel1.Controls.Add(this.label3, 5, 0);
            this.tableLayoutPanel1.Controls.Add(this.label4, 5, 1);
            this.tableLayoutPanel1.Controls.Add(this.comboBoxDelimeter, 6, 1);
            this.tableLayoutPanel1.Controls.Add(this.label6, 5, 2);
            this.tableLayoutPanel1.Controls.Add(this.comboBoxStile, 6, 2);
            this.tableLayoutPanel1.Controls.Add(this.label5, 5, 3);
            this.tableLayoutPanel1.Controls.Add(this.numericUpDownColVal, 6, 3);
            this.tableLayoutPanel1.Controls.Add(this.label7, 5, 6);
            this.tableLayoutPanel1.Controls.Add(this.button2, 6, 6);
            this.tableLayoutPanel1.Controls.Add(this.textBoxDirectory, 5, 7);
            this.tableLayoutPanel1.Controls.Add(this.label8, 5, 8);
            this.tableLayoutPanel1.Controls.Add(this.textBoxTableName, 6, 8);
            this.tableLayoutPanel1.Controls.Add(this.label10, 5, 4);
            this.tableLayoutPanel1.Controls.Add(this.label9, 5, 5);
            this.tableLayoutPanel1.Controls.Add(this.numericUpDownDecPlaces, 6, 4);
            this.tableLayoutPanel1.Controls.Add(this.comboBoxSeparator, 6, 5);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 13;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 46F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 46F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(776, 335);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Location = new System.Drawing.Point(369, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(126, 32);
            this.label2.TabIndex = 3;
            this.label2.Text = "Сut-off";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // textBoxValue
            // 
            this.textBoxValue.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxValue.Location = new System.Drawing.Point(369, 214);
            this.textBoxValue.Name = "textBoxValue";
            this.textBoxValue.Size = new System.Drawing.Size(126, 26);
            this.textBoxValue.TabIndex = 14;
            this.textBoxValue.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBoxValue_KeyPress);
            // 
            // listBoxSource
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.listBoxSource, 2);
            this.listBoxSource.ContextMenuStrip = this.contextMenuStripSource;
            this.listBoxSource.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxSource.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.listBoxSource.FormattingEnabled = true;
            this.listBoxSource.HorizontalScrollbar = true;
            this.listBoxSource.ItemHeight = 20;
            this.listBoxSource.Location = new System.Drawing.Point(3, 35);
            this.listBoxSource.Name = "listBoxSource";
            this.tableLayoutPanel1.SetRowSpan(this.listBoxSource, 11);
            this.listBoxSource.ScrollAlwaysVisible = true;
            this.listBoxSource.Size = new System.Drawing.Size(352, 251);
            this.listBoxSource.TabIndex = 0;
            // 
            // contextMenuStripSource
            // 
            this.contextMenuStripSource.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addFileToolStripMenuItem,
            this.deleteFileToolStripMenuItem,
            this.deleteAllToolStripMenuItem});
            this.contextMenuStripSource.Name = "contextMenuStripSource";
            this.contextMenuStripSource.Size = new System.Drawing.Size(118, 70);
            // 
            // addFileToolStripMenuItem
            // 
            this.addFileToolStripMenuItem.Name = "addFileToolStripMenuItem";
            this.addFileToolStripMenuItem.Size = new System.Drawing.Size(117, 22);
            this.addFileToolStripMenuItem.Text = "Add File";
            this.addFileToolStripMenuItem.Click += new System.EventHandler(this.addFileToolStripMenuItem_Click);
            // 
            // deleteFileToolStripMenuItem
            // 
            this.deleteFileToolStripMenuItem.Name = "deleteFileToolStripMenuItem";
            this.deleteFileToolStripMenuItem.Size = new System.Drawing.Size(117, 22);
            this.deleteFileToolStripMenuItem.Text = "Delete";
            this.deleteFileToolStripMenuItem.Click += new System.EventHandler(this.deleteFileToolStripMenuItem_Click);
            // 
            // deleteAllToolStripMenuItem
            // 
            this.deleteAllToolStripMenuItem.Name = "deleteAllToolStripMenuItem";
            this.deleteAllToolStripMenuItem.Size = new System.Drawing.Size(117, 22);
            this.deleteAllToolStripMenuItem.Text = "Clear";
            this.deleteAllToolStripMenuItem.Click += new System.EventHandler(this.deleteAllToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.label1, 2);
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(352, 32);
            this.label1.TabIndex = 15;
            this.label1.Text = "Source File";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // buttonUp
            // 
            this.buttonUp.Dock = System.Windows.Forms.DockStyle.Fill;
            this.buttonUp.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonUp.Location = new System.Drawing.Point(3, 292);
            this.buttonUp.Name = "buttonUp";
            this.buttonUp.Size = new System.Drawing.Size(172, 40);
            this.buttonUp.TabIndex = 17;
            this.buttonUp.Text = "Up";
            this.buttonUp.UseVisualStyleBackColor = true;
            this.buttonUp.Click += new System.EventHandler(this.buttonUp_Click);
            // 
            // buttonDown
            // 
            this.buttonDown.Dock = System.Windows.Forms.DockStyle.Fill;
            this.buttonDown.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonDown.Location = new System.Drawing.Point(181, 292);
            this.buttonDown.Name = "buttonDown";
            this.buttonDown.Size = new System.Drawing.Size(174, 40);
            this.buttonDown.TabIndex = 18;
            this.buttonDown.Text = "Down";
            this.buttonDown.UseVisualStyleBackColor = true;
            this.buttonDown.Click += new System.EventHandler(this.buttonDown_Click);
            // 
            // buttonUpVal
            // 
            this.buttonUpVal.Dock = System.Windows.Forms.DockStyle.Fill;
            this.buttonUpVal.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonUpVal.Location = new System.Drawing.Point(369, 246);
            this.buttonUpVal.Name = "buttonUpVal";
            this.buttonUpVal.Size = new System.Drawing.Size(126, 40);
            this.buttonUpVal.TabIndex = 11;
            this.buttonUpVal.Text = "Up";
            this.buttonUpVal.UseVisualStyleBackColor = true;
            this.buttonUpVal.Click += new System.EventHandler(this.buttonUpVal_Click);
            // 
            // buttonDownVal
            // 
            this.buttonDownVal.Dock = System.Windows.Forms.DockStyle.Fill;
            this.buttonDownVal.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonDownVal.Location = new System.Drawing.Point(369, 292);
            this.buttonDownVal.Name = "buttonDownVal";
            this.buttonDownVal.Size = new System.Drawing.Size(126, 40);
            this.buttonDownVal.TabIndex = 19;
            this.buttonDownVal.Text = "Down";
            this.buttonDownVal.UseVisualStyleBackColor = true;
            this.buttonDownVal.Click += new System.EventHandler(this.buttonDownVal_Click);
            // 
            // listBoxValue
            // 
            this.listBoxValue.BackColor = System.Drawing.SystemColors.Control;
            this.listBoxValue.ContextMenuStrip = this.contextMenuStripValue;
            this.listBoxValue.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.listBoxValue.ItemHeight = 20;
            this.listBoxValue.Location = new System.Drawing.Point(369, 35);
            this.listBoxValue.Name = "listBoxValue";
            this.listBoxValue.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.tableLayoutPanel1.SetRowSpan(this.listBoxValue, 9);
            this.listBoxValue.Size = new System.Drawing.Size(126, 173);
            this.listBoxValue.TabIndex = 20;
            // 
            // contextMenuStripValue
            // 
            this.contextMenuStripValue.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.deleteToolStripMenuItem,
            this.clearToolStripMenuItem});
            this.contextMenuStripValue.Name = "contextMenuStripValue";
            this.contextMenuStripValue.Size = new System.Drawing.Size(108, 48);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            this.deleteToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.deleteToolStripMenuItem.Text = "Delete";
            this.deleteToolStripMenuItem.Click += new System.EventHandler(this.deleteToolStripMenuItem_Click);
            // 
            // clearToolStripMenuItem
            // 
            this.clearToolStripMenuItem.Name = "clearToolStripMenuItem";
            this.clearToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.clearToolStripMenuItem.Text = "Clear";
            this.clearToolStripMenuItem.Click += new System.EventHandler(this.clearToolStripMenuItem_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.label3, 2);
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Location = new System.Drawing.Point(508, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(265, 32);
            this.label3.TabIndex = 21;
            this.label3.Text = "General Settings";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Location = new System.Drawing.Point(508, 32);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(94, 32);
            this.label4.TabIndex = 22;
            this.label4.Text = "Delimeter";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // comboBoxDelimeter
            // 
            this.comboBoxDelimeter.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBoxDelimeter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBoxDelimeter.FormattingEnabled = true;
            this.comboBoxDelimeter.Items.AddRange(new object[] {
            "TAB",
            "Comma",
            "Point",
            "Semicolon",
            "Space"});
            this.comboBoxDelimeter.Location = new System.Drawing.Point(608, 35);
            this.comboBoxDelimeter.Name = "comboBoxDelimeter";
            this.comboBoxDelimeter.Size = new System.Drawing.Size(165, 24);
            this.comboBoxDelimeter.TabIndex = 23;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label6.Location = new System.Drawing.Point(508, 64);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(94, 32);
            this.label6.TabIndex = 26;
            this.label6.Text = "Writing stile";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // comboBoxStile
            // 
            this.comboBoxStile.DisplayMember = "1";
            this.comboBoxStile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBoxStile.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBoxStile.Items.AddRange(new object[] {
            "Source per Row",
            "Source per Column"});
            this.comboBoxStile.Location = new System.Drawing.Point(608, 67);
            this.comboBoxStile.Name = "comboBoxStile";
            this.comboBoxStile.Size = new System.Drawing.Size(165, 24);
            this.comboBoxStile.TabIndex = 27;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label5.Location = new System.Drawing.Point(508, 96);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(94, 32);
            this.label5.TabIndex = 24;
            this.label5.Text = "Value\'s Column ";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // numericUpDownColVal
            // 
            this.numericUpDownColVal.Dock = System.Windows.Forms.DockStyle.Fill;
            this.numericUpDownColVal.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.numericUpDownColVal.Location = new System.Drawing.Point(608, 99);
            this.numericUpDownColVal.Minimum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.numericUpDownColVal.Name = "numericUpDownColVal";
            this.numericUpDownColVal.Size = new System.Drawing.Size(165, 22);
            this.numericUpDownColVal.TabIndex = 25;
            this.numericUpDownColVal.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label7.Location = new System.Drawing.Point(508, 192);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(94, 32);
            this.label7.TabIndex = 28;
            this.label7.Text = "Folder";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // button2
            // 
            this.button2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button2.Location = new System.Drawing.Point(608, 195);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(165, 26);
            this.button2.TabIndex = 29;
            this.button2.Text = "Change";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBoxDirectory
            // 
            this.textBoxDirectory.BackColor = System.Drawing.SystemColors.Control;
            this.tableLayoutPanel1.SetColumnSpan(this.textBoxDirectory, 2);
            this.textBoxDirectory.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.textBoxDirectory.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxDirectory.Location = new System.Drawing.Point(508, 231);
            this.textBoxDirectory.Name = "textBoxDirectory";
            this.textBoxDirectory.Size = new System.Drawing.Size(265, 22);
            this.textBoxDirectory.TabIndex = 30;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label8.Location = new System.Drawing.Point(508, 256);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(94, 32);
            this.label8.TabIndex = 31;
            this.label8.Text = "Table Name";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // textBoxTableName
            // 
            this.textBoxTableName.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxTableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxTableName.Location = new System.Drawing.Point(608, 259);
            this.textBoxTableName.Name = "textBoxTableName";
            this.textBoxTableName.Size = new System.Drawing.Size(165, 22);
            this.textBoxTableName.TabIndex = 32;
            this.textBoxTableName.Text = "Table";
            this.textBoxTableName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.PointPressed);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label10.Location = new System.Drawing.Point(508, 128);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(94, 32);
            this.label10.TabIndex = 34;
            this.label10.Text = "Decimal Places";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label9.Location = new System.Drawing.Point(508, 160);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(94, 32);
            this.label9.TabIndex = 33;
            this.label9.Text = "Decimal Separator";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // numericUpDownDecPlaces
            // 
            this.numericUpDownDecPlaces.Dock = System.Windows.Forms.DockStyle.Fill;
            this.numericUpDownDecPlaces.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.numericUpDownDecPlaces.Location = new System.Drawing.Point(608, 131);
            this.numericUpDownDecPlaces.Name = "numericUpDownDecPlaces";
            this.numericUpDownDecPlaces.Size = new System.Drawing.Size(165, 22);
            this.numericUpDownDecPlaces.TabIndex = 35;
            // 
            // comboBoxSeparator
            // 
            this.comboBoxSeparator.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBoxSeparator.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBoxSeparator.FormattingEnabled = true;
            this.comboBoxSeparator.Items.AddRange(new object[] {
            "Comma",
            "Point"});
            this.comboBoxSeparator.Location = new System.Drawing.Point(608, 163);
            this.comboBoxSeparator.Name = "comboBoxSeparator";
            this.comboBoxSeparator.Size = new System.Drawing.Size(165, 24);
            this.comboBoxSeparator.TabIndex = 36;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat|All Files (*.*)|*.*";
            this.openFileDialog1.Multiselect = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(776, 335);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "TableDataExtractor v1.1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.contextMenuStripSource.ResumeLayout(false);
            this.contextMenuStripValue.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownColVal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownDecPlaces)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonUpVal;
        private System.Windows.Forms.TextBox textBoxValue;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox listBoxSource;
        private System.Windows.Forms.Button buttonUp;
        private System.Windows.Forms.Button buttonDown;
        private System.Windows.Forms.Button buttonDownVal;
        private System.Windows.Forms.ListBox listBoxValue;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripSource;
        private System.Windows.Forms.ToolStripMenuItem addFileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deleteFileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deleteAllToolStripMenuItem;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripValue;
        private System.Windows.Forms.ToolStripMenuItem deleteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem clearToolStripMenuItem;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox comboBoxDelimeter;
        private System.Windows.Forms.NumericUpDown numericUpDownColVal;
        private System.Windows.Forms.ComboBox comboBoxStile;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox textBoxDirectory;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox textBoxTableName;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.NumericUpDown numericUpDownDecPlaces;
        private System.Windows.Forms.ComboBox comboBoxSeparator;
    }
}

