namespace Fiverr_spreadsheet_project_1
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            label1 = new Label();
            csvTextBox = new TextBox();
            csvBrowseButton = new Button();
            label2 = new Label();
            excelTextBox = new TextBox();
            excelBrowseButton = new Button();
            convertButton = new Button();
            openCSVFileDialog = new OpenFileDialog();
            openExcelFileDialog = new OpenFileDialog();
            saveExcelFileDialog = new SaveFileDialog();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(67, 28);
            label1.Name = "label1";
            label1.Size = new Size(121, 15);
            label1.TabIndex = 0;
            label1.Text = "Hour Report Location";
            // 
            // csvTextBox
            // 
            csvTextBox.Location = new Point(12, 46);
            csvTextBox.Name = "csvTextBox";
            csvTextBox.ReadOnly = true;
            csvTextBox.Size = new Size(206, 23);
            csvTextBox.TabIndex = 1;
            // 
            // csvBrowseButton
            // 
            csvBrowseButton.Location = new Point(82, 75);
            csvBrowseButton.Name = "csvBrowseButton";
            csvBrowseButton.Size = new Size(75, 23);
            csvBrowseButton.TabIndex = 2;
            csvBrowseButton.Text = "Browse...";
            csvBrowseButton.UseVisualStyleBackColor = true;
            csvBrowseButton.Click += csvBrowseButton_Click;
            // 
            // label2
            // 
            label2.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label2.AutoSize = true;
            label2.Location = new Point(327, 28);
            label2.Name = "label2";
            label2.Size = new Size(99, 15);
            label2.TabIndex = 3;
            label2.Text = "Analysis Location";
            // 
            // excelTextBox
            // 
            excelTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            excelTextBox.Location = new Point(280, 46);
            excelTextBox.Name = "excelTextBox";
            excelTextBox.ReadOnly = true;
            excelTextBox.Size = new Size(192, 23);
            excelTextBox.TabIndex = 4;
            // 
            // excelBrowseButton
            // 
            excelBrowseButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            excelBrowseButton.Location = new Point(339, 75);
            excelBrowseButton.Name = "excelBrowseButton";
            excelBrowseButton.Size = new Size(75, 23);
            excelBrowseButton.TabIndex = 5;
            excelBrowseButton.Text = "Browse...";
            excelBrowseButton.UseVisualStyleBackColor = true;
            excelBrowseButton.Click += excelBrowseButton_Click;
            // 
            // convertButton
            // 
            convertButton.Anchor = AnchorStyles.Bottom;
            convertButton.Location = new Point(205, 348);
            convertButton.Name = "convertButton";
            convertButton.Size = new Size(75, 23);
            convertButton.TabIndex = 8;
            convertButton.Text = "Convert";
            convertButton.UseVisualStyleBackColor = true;
            convertButton.Click += convertButton_Click;
            // 
            // openCSVFileDialog
            // 
            openCSVFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            openCSVFileDialog.Title = "Open \".csv\" File";
            // 
            // openExcelFileDialog
            // 
            openExcelFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openExcelFileDialog.Title = "Choose Excel File";
            // 
            // saveExcelFileDialog
            // 
            saveExcelFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveExcelFileDialog.Title = "Save As Excel File";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(484, 461);
            Controls.Add(convertButton);
            Controls.Add(excelBrowseButton);
            Controls.Add(excelTextBox);
            Controls.Add(label2);
            Controls.Add(csvBrowseButton);
            Controls.Add(csvTextBox);
            Controls.Add(label1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Fiverr Project";
            Load += Form1_Load;
            Resize += Form1_Resize;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private TextBox csvTextBox;
        private Button csvBrowseButton;
        private Label label2;
        private TextBox excelTextBox;
        private Button excelBrowseButton;
        private Button convertButton;
        private OpenFileDialog openCSVFileDialog;
        private OpenFileDialog openExcelFileDialog;
        private SaveFileDialog saveExcelFileDialog;
    }
}
