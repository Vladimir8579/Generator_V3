﻿namespace Generator_V3
{
    partial class Generator
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.Generation = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.Exit = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.PanelWord = new System.Windows.Forms.Panel();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.checkedListBox2 = new System.Windows.Forms.CheckedListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SelectWord = new System.Windows.Forms.Button();
            this.textBox1SelectWord = new System.Windows.Forms.TextBox();
            this.PanelExcel = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.SelectExcel = new System.Windows.Forms.Button();
            this.textBoxSelectExcel = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.toolStrip2 = new System.Windows.Forms.ToolStrip();
            this.toolStripLabel2 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripComboBox2 = new System.Windows.Forms.ToolStripComboBox();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripComboBox1 = new System.Windows.Forms.ToolStripComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.PanelPatchSave = new System.Windows.Forms.Panel();
            this.CheckBoxSaveToPdf = new System.Windows.Forms.CheckBox();
            this.label10 = new System.Windows.Forms.Label();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.label9 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.SelectPathSave = new System.Windows.Forms.Button();
            this.textBoxSelectPathSave = new System.Windows.Forms.TextBox();
            this.PanelWord.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.PanelExcel.SuspendLayout();
            this.toolStrip2.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.PanelPatchSave.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.SuspendLayout();
            // 
            // Generation
            // 
            this.Generation.Enabled = false;
            this.Generation.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Generation.Location = new System.Drawing.Point(8, 627);
            this.Generation.Margin = new System.Windows.Forms.Padding(4);
            this.Generation.Name = "Generation";
            this.Generation.Size = new System.Drawing.Size(160, 50);
            this.Generation.TabIndex = 13;
            this.Generation.Text = "Сгенерировать";
            this.Generation.UseVisualStyleBackColor = true;
            this.Generation.Click += new System.EventHandler(this.Generation_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "| *.xlsx";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            this.openFileDialog2.Filter = "| *.docx";
            // 
            // Exit
            // 
            this.Exit.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Exit.Location = new System.Drawing.Point(667, 627);
            this.Exit.Margin = new System.Windows.Forms.Padding(4);
            this.Exit.Name = "Exit";
            this.Exit.Size = new System.Drawing.Size(160, 50);
            this.Exit.TabIndex = 30;
            this.Exit.Text = "Выход";
            this.Exit.UseVisualStyleBackColor = true;
            this.Exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // PanelWord
            // 
            this.PanelWord.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.PanelWord.Controls.Add(this.dataGridView2);
            this.PanelWord.Controls.Add(this.dataGridView1);
            this.PanelWord.Controls.Add(this.label2);
            this.PanelWord.Controls.Add(this.checkedListBox2);
            this.PanelWord.Controls.Add(this.label4);
            this.PanelWord.Controls.Add(this.checkedListBox1);
            this.PanelWord.Controls.Add(this.label1);
            this.PanelWord.Controls.Add(this.SelectWord);
            this.PanelWord.Controls.Add(this.textBox1SelectWord);
            this.PanelWord.Location = new System.Drawing.Point(8, 3);
            this.PanelWord.Name = "PanelWord";
            this.PanelWord.Size = new System.Drawing.Size(819, 324);
            this.PanelWord.TabIndex = 35;
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(707, 107);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(78, 24);
            this.dataGridView2.TabIndex = 36;
            this.dataGridView2.Visible = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(707, 42);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(78, 26);
            this.dataGridView1.TabIndex = 35;
            this.dataGridView1.Visible = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label2.Location = new System.Drawing.Point(11, 8);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.MaximumSize = new System.Drawing.Size(240, 24);
            this.label2.MinimumSize = new System.Drawing.Size(240, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(240, 24);
            this.label2.TabIndex = 34;
            this.label2.Text = "Выберите шаблон Word";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // checkedListBox2
            // 
            this.checkedListBox2.CheckOnClick = true;
            this.checkedListBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedListBox2.FormattingEnabled = true;
            this.checkedListBox2.Location = new System.Drawing.Point(259, 222);
            this.checkedListBox2.Margin = new System.Windows.Forms.Padding(4);
            this.checkedListBox2.Name = "checkedListBox2";
            this.checkedListBox2.Size = new System.Drawing.Size(428, 89);
            this.checkedListBox2.TabIndex = 33;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label4.Location = new System.Drawing.Point(11, 222);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.MaximumSize = new System.Drawing.Size(240, 89);
            this.label4.MinimumSize = new System.Drawing.Size(240, 89);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(240, 89);
            this.label4.TabIndex = 32;
            this.label4.Text = "Выберите таблицы, в которых\r\nнужно удалить пустые строки";
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.CheckOnClick = true;
            this.checkedListBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.HorizontalScrollbar = true;
            this.checkedListBox1.Location = new System.Drawing.Point(259, 40);
            this.checkedListBox1.Margin = new System.Windows.Forms.Padding(4);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(428, 174);
            this.checkedListBox1.TabIndex = 31;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label1.Location = new System.Drawing.Point(11, 40);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.MaximumSize = new System.Drawing.Size(240, 174);
            this.label1.MinimumSize = new System.Drawing.Size(240, 174);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(240, 174);
            this.label1.TabIndex = 30;
            this.label1.Text = "Выберите блоки текста,\r\nкоторые необходимо оставить";
            // 
            // SelectWord
            // 
            this.SelectWord.Location = new System.Drawing.Point(707, 8);
            this.SelectWord.Margin = new System.Windows.Forms.Padding(4);
            this.SelectWord.MaximumSize = new System.Drawing.Size(100, 24);
            this.SelectWord.MinimumSize = new System.Drawing.Size(100, 24);
            this.SelectWord.Name = "SelectWord";
            this.SelectWord.Size = new System.Drawing.Size(100, 24);
            this.SelectWord.TabIndex = 29;
            this.SelectWord.Text = "Обзор";
            this.SelectWord.UseVisualStyleBackColor = true;
            this.SelectWord.Click += new System.EventHandler(this.SelectWord_Click);
            // 
            // textBox1SelectWord
            // 
            this.textBox1SelectWord.Location = new System.Drawing.Point(259, 8);
            this.textBox1SelectWord.Margin = new System.Windows.Forms.Padding(4);
            this.textBox1SelectWord.MaximumSize = new System.Drawing.Size(428, 24);
            this.textBox1SelectWord.MinimumSize = new System.Drawing.Size(428, 24);
            this.textBox1SelectWord.Name = "textBox1SelectWord";
            this.textBox1SelectWord.Size = new System.Drawing.Size(428, 22);
            this.textBox1SelectWord.TabIndex = 28;
            this.textBox1SelectWord.TextChanged += new System.EventHandler(this.GenerationButtonCheked);
            this.textBox1SelectWord.DoubleClick += new System.EventHandler(this.SelectWord_Click);
            // 
            // PanelExcel
            // 
            this.PanelExcel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.PanelExcel.Controls.Add(this.label5);
            this.PanelExcel.Controls.Add(this.label3);
            this.PanelExcel.Controls.Add(this.SelectExcel);
            this.PanelExcel.Controls.Add(this.textBoxSelectExcel);
            this.PanelExcel.Controls.Add(this.label8);
            this.PanelExcel.Controls.Add(this.comboBox1);
            this.PanelExcel.Controls.Add(this.toolStrip2);
            this.PanelExcel.Controls.Add(this.toolStrip1);
            this.PanelExcel.Controls.Add(this.label6);
            this.PanelExcel.Location = new System.Drawing.Point(8, 338);
            this.PanelExcel.Name = "PanelExcel";
            this.PanelExcel.Size = new System.Drawing.Size(819, 147);
            this.PanelExcel.TabIndex = 36;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label5.Location = new System.Drawing.Point(11, 46);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.MaximumSize = new System.Drawing.Size(240, 24);
            this.label5.MinimumSize = new System.Drawing.Size(240, 24);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(240, 24);
            this.label5.TabIndex = 33;
            this.label5.Text = "Выберите лист с переменными";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label3.Location = new System.Drawing.Point(11, 10);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.MaximumSize = new System.Drawing.Size(240, 24);
            this.label3.MinimumSize = new System.Drawing.Size(240, 24);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(240, 24);
            this.label3.TabIndex = 32;
            this.label3.Text = "Выберите таблицу Excel";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // SelectExcel
            // 
            this.SelectExcel.Location = new System.Drawing.Point(707, 9);
            this.SelectExcel.Margin = new System.Windows.Forms.Padding(4);
            this.SelectExcel.MaximumSize = new System.Drawing.Size(100, 24);
            this.SelectExcel.MinimumSize = new System.Drawing.Size(100, 24);
            this.SelectExcel.Name = "SelectExcel";
            this.SelectExcel.Size = new System.Drawing.Size(100, 24);
            this.SelectExcel.TabIndex = 31;
            this.SelectExcel.Text = "Обзор";
            this.SelectExcel.UseVisualStyleBackColor = true;
            this.SelectExcel.Click += new System.EventHandler(this.SelectExcel_Click);
            // 
            // textBoxSelectExcel
            // 
            this.textBoxSelectExcel.Location = new System.Drawing.Point(259, 10);
            this.textBoxSelectExcel.Margin = new System.Windows.Forms.Padding(4);
            this.textBoxSelectExcel.MaximumSize = new System.Drawing.Size(428, 24);
            this.textBoxSelectExcel.MinimumSize = new System.Drawing.Size(428, 24);
            this.textBoxSelectExcel.Name = "textBoxSelectExcel";
            this.textBoxSelectExcel.Size = new System.Drawing.Size(428, 22);
            this.textBoxSelectExcel.TabIndex = 30;
            this.textBoxSelectExcel.TextChanged += new System.EventHandler(this.GenerationButtonCheked);
            this.textBoxSelectExcel.DoubleClick += new System.EventHandler(this.SelectExcel_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label8.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label8.Location = new System.Drawing.Point(11, 113);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.MaximumSize = new System.Drawing.Size(240, 24);
            this.label8.MinimumSize = new System.Drawing.Size(240, 24);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(240, 24);
            this.label8.TabIndex = 38;
            this.label8.Text = "Выберите поле с именем файлов";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(259, 113);
            this.comboBox1.MaximumSize = new System.Drawing.Size(428, 0);
            this.comboBox1.MinimumSize = new System.Drawing.Size(428, 0);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(428, 24);
            this.comboBox1.TabIndex = 37;
            // 
            // toolStrip2
            // 
            this.toolStrip2.Dock = System.Windows.Forms.DockStyle.None;
            this.toolStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripLabel2,
            this.toolStripComboBox2});
            this.toolStrip2.Location = new System.Drawing.Point(259, 80);
            this.toolStrip2.MaximumSize = new System.Drawing.Size(428, 24);
            this.toolStrip2.MinimumSize = new System.Drawing.Size(428, 24);
            this.toolStrip2.Name = "toolStrip2";
            this.toolStrip2.Size = new System.Drawing.Size(428, 24);
            this.toolStrip2.TabIndex = 36;
            this.toolStrip2.Text = "toolStrip2";
            // 
            // toolStripLabel2
            // 
            this.toolStripLabel2.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.toolStripLabel2.Name = "toolStripLabel2";
            this.toolStripLabel2.Size = new System.Drawing.Size(43, 21);
            this.toolStripLabel2.Text = "Лист";
            // 
            // toolStripComboBox2
            // 
            this.toolStripComboBox2.AutoSize = false;
            this.toolStripComboBox2.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.toolStripComboBox2.Name = "toolStripComboBox2";
            this.toolStripComboBox2.Size = new System.Drawing.Size(345, 27);
            this.toolStripComboBox2.SelectedIndexChanged += new System.EventHandler(this.ToolStripComboBox2_SelectedIndexChanged);
            // 
            // toolStrip1
            // 
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripLabel1,
            this.toolStripComboBox1});
            this.toolStrip1.Location = new System.Drawing.Point(259, 46);
            this.toolStrip1.MaximumSize = new System.Drawing.Size(428, 24);
            this.toolStrip1.MinimumSize = new System.Drawing.Size(428, 24);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(428, 24);
            this.toolStrip1.TabIndex = 35;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(43, 21);
            this.toolStripLabel1.Text = "Лист";
            // 
            // toolStripComboBox1
            // 
            this.toolStripComboBox1.AutoSize = false;
            this.toolStripComboBox1.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.toolStripComboBox1.Name = "toolStripComboBox1";
            this.toolStripComboBox1.Size = new System.Drawing.Size(345, 27);
            this.toolStripComboBox1.SelectedIndexChanged += new System.EventHandler(this.ToolStripComboBox1_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label6.Location = new System.Drawing.Point(11, 80);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.MaximumSize = new System.Drawing.Size(240, 24);
            this.label6.MinimumSize = new System.Drawing.Size(240, 24);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(240, 24);
            this.label6.TabIndex = 34;
            this.label6.Text = "Выберите лист со списками";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // PanelPatchSave
            // 
            this.PanelPatchSave.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.PanelPatchSave.Controls.Add(this.CheckBoxSaveToPdf);
            this.PanelPatchSave.Controls.Add(this.label10);
            this.PanelPatchSave.Controls.Add(this.numericUpDown1);
            this.PanelPatchSave.Controls.Add(this.label9);
            this.PanelPatchSave.Controls.Add(this.label7);
            this.PanelPatchSave.Controls.Add(this.SelectPathSave);
            this.PanelPatchSave.Controls.Add(this.textBoxSelectPathSave);
            this.PanelPatchSave.Location = new System.Drawing.Point(8, 497);
            this.PanelPatchSave.Name = "PanelPatchSave";
            this.PanelPatchSave.Size = new System.Drawing.Size(819, 123);
            this.PanelPatchSave.TabIndex = 37;
            // 
            // CheckBoxSaveToPdf
            // 
            this.CheckBoxSaveToPdf.AutoSize = true;
            this.CheckBoxSaveToPdf.Location = new System.Drawing.Point(259, 84);
            this.CheckBoxSaveToPdf.MaximumSize = new System.Drawing.Size(260, 24);
            this.CheckBoxSaveToPdf.MinimumSize = new System.Drawing.Size(260, 24);
            this.CheckBoxSaveToPdf.Name = "CheckBoxSaveToPdf";
            this.CheckBoxSaveToPdf.Size = new System.Drawing.Size(260, 24);
            this.CheckBoxSaveToPdf.TabIndex = 41;
            this.CheckBoxSaveToPdf.Text = "Дополнительно сохранить в PDF";
            this.CheckBoxSaveToPdf.UseVisualStyleBackColor = true;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label10.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label10.Location = new System.Drawing.Point(11, 84);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.MaximumSize = new System.Drawing.Size(240, 24);
            this.label10.MinimumSize = new System.Drawing.Size(240, 24);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(240, 24);
            this.label10.TabIndex = 40;
            this.label10.Text = "Сохранить в PDF";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label10.Visible = false;
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Location = new System.Drawing.Point(259, 49);
            this.numericUpDown1.Maximum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.numericUpDown1.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(428, 22);
            this.numericUpDown1.TabIndex = 39;
            this.numericUpDown1.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label9.Location = new System.Drawing.Point(11, 47);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.MaximumSize = new System.Drawing.Size(240, 24);
            this.label9.MinimumSize = new System.Drawing.Size(240, 24);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(240, 24);
            this.label9.TabIndex = 38;
            this.label9.Text = "Выберите, количество экз-ов";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label7.Location = new System.Drawing.Point(11, 10);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.MaximumSize = new System.Drawing.Size(240, 24);
            this.label7.MinimumSize = new System.Drawing.Size(240, 24);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(240, 24);
            this.label7.TabIndex = 37;
            this.label7.Text = "Выберите, куда сохранить файл";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // SelectPathSave
            // 
            this.SelectPathSave.Location = new System.Drawing.Point(707, 9);
            this.SelectPathSave.Margin = new System.Windows.Forms.Padding(4);
            this.SelectPathSave.MaximumSize = new System.Drawing.Size(100, 24);
            this.SelectPathSave.MinimumSize = new System.Drawing.Size(100, 24);
            this.SelectPathSave.Name = "SelectPathSave";
            this.SelectPathSave.Size = new System.Drawing.Size(100, 24);
            this.SelectPathSave.TabIndex = 36;
            this.SelectPathSave.Text = "Обзор";
            this.SelectPathSave.UseVisualStyleBackColor = true;
            this.SelectPathSave.Click += new System.EventHandler(this.SelectPathSave_Click);
            // 
            // textBoxSelectPathSave
            // 
            this.textBoxSelectPathSave.Location = new System.Drawing.Point(259, 9);
            this.textBoxSelectPathSave.Margin = new System.Windows.Forms.Padding(4);
            this.textBoxSelectPathSave.MaximumSize = new System.Drawing.Size(428, 24);
            this.textBoxSelectPathSave.MinimumSize = new System.Drawing.Size(428, 24);
            this.textBoxSelectPathSave.Name = "textBoxSelectPathSave";
            this.textBoxSelectPathSave.Size = new System.Drawing.Size(428, 22);
            this.textBoxSelectPathSave.TabIndex = 35;
            this.textBoxSelectPathSave.TextChanged += new System.EventHandler(this.GenerationButtonCheked);
            this.textBoxSelectPathSave.DoubleClick += new System.EventHandler(this.SelectPathSave_Click);
            // 
            // Generator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(835, 683);
            this.Controls.Add(this.PanelPatchSave);
            this.Controls.Add(this.PanelExcel);
            this.Controls.Add(this.PanelWord);
            this.Controls.Add(this.Exit);
            this.Controls.Add(this.Generation);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "Generator";
            this.Text = "Generator_V3";
            this.PanelWord.ResumeLayout(false);
            this.PanelWord.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.PanelExcel.ResumeLayout(false);
            this.PanelExcel.PerformLayout();
            this.toolStrip2.ResumeLayout(false);
            this.toolStrip2.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.PanelPatchSave.ResumeLayout(false);
            this.PanelPatchSave.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button Generation;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.Button Exit;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Panel PanelWord;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckedListBox checkedListBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button SelectWord;
        private System.Windows.Forms.TextBox textBox1SelectWord;
        private System.Windows.Forms.Panel PanelExcel;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button SelectExcel;
        private System.Windows.Forms.TextBox textBoxSelectExcel;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ToolStrip toolStrip2;
        private System.Windows.Forms.ToolStripLabel toolStripLabel2;
        private System.Windows.Forms.ToolStripComboBox toolStripComboBox2;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripComboBox toolStripComboBox1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Panel PanelPatchSave;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button SelectPathSave;
        private System.Windows.Forms.TextBox textBoxSelectPathSave;
        private System.Windows.Forms.CheckBox CheckBoxSaveToPdf;
    }
}

