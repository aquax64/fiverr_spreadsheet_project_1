using Fiverr_spreadsheet_project_1.Libraries;
using System.Windows.Forms;

namespace Fiverr_spreadsheet_project_1
{
    public partial class Form1 : Form
    {
        private bool appending = true;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Load Config Files
            Loader.LoadConfig();
            Loader.LoadOtherConfigs();
        }

        private bool isResizing = false;

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (isResizing)
                return;

            isResizing = true;

            // Enforce Min 500 by 500
            if (this.Height < 500)
            {
                this.Height = 500;
            }
            if (this.Width < 500)
            {
                this.Width = 500;
            }

            isResizing = false;
        }

        private void csvBrowseButton_Click(object sender, EventArgs e)
        {
            openCSVFileDialog.InitialDirectory = Environment.CurrentDirectory;

            if (openCSVFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openCSVFileDialog.FileName;

                csvTextBox.Text = filePath;
            }
        }

        private void excelBrowseButton_Click(object sender, EventArgs e)
        {

            if (appending)
            {

                openExcelFileDialog.InitialDirectory = Environment.CurrentDirectory;

                if (openExcelFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openExcelFileDialog.FileName;

                    excelTextBox.Text = filePath;
                }
            } else
            {

                saveExcelFileDialog.InitialDirectory = Environment.CurrentDirectory;

                if (saveExcelFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveExcelFileDialog.FileName;

                    excelTextBox.Text = filePath;
                }
            }
        }

        private void convertButton_Click(object sender, EventArgs e)
        {
            DataHandler handler = new DataHandler();

            handler.ConvertCsvToExcel(csvTextBox.Text, excelTextBox.Text, appending);
        }
    }
}
