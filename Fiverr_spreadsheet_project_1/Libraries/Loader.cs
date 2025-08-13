using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fiverr_spreadsheet_project_1.Libraries
{
    public static class Loader
    {
        static string configPath = string.Empty;
        static string configText = string.Empty;

        public static Dictionary<string, string> configPairs = new Dictionary<string, string>();

        public static int ADMIN_START = -1;
        public static int TITLEROW = -1;

        public static void LoadConfig()
        {
            configPath = Path.Combine(Environment.CurrentDirectory, "config.txt"); // Or use Application.StartupPath
            configText = File.ReadAllText(configPath);

            string[] unfilteredPairs = configText.Split(',');

            foreach (var pair in unfilteredPairs)
            {
                string? editedPair;

                editedPair = pair.Replace("\n", ""); // Remove new lines
                editedPair = editedPair.Replace("\r", "");
                editedPair = editedPair.Replace(": ", "");

                string csvString = "";
                string excelString = "";

                // Find csv word
                int lastIndex = -1;
                for (int i = 1; i < editedPair.Length; i++)
                {
                    if (editedPair[i] == '\"')
                    {
                        lastIndex = i;
                        break;
                    }

                    csvString += editedPair[i];
                }

                if (lastIndex == -1)
                {
                    MessageBox.Show("ERROR: Config formatted incorrectly");
                }

                // Find excel word
                for (int i = lastIndex + 2; i < editedPair.Length; i++) // Skip past from " to : to " to the first letter
                {
                    if (editedPair[i] == '\"')
                        continue;

                    excelString += editedPair[i];
                }

                // Add to dictionary
                configPairs.Add(csvString, excelString);
            }

            Console.WriteLine("Finished loading config");
        }

        public static void LoadOtherConfigs()
        {
            configPath = Path.Combine(Environment.CurrentDirectory, "csvstart.txt");
            configText = File.ReadAllText(configPath);

            string[] configPairs = configText.Split('\n');

            try
            {
                foreach (string configPair in configPairs)
                {
                    string[] splits = configPair.Split('=');

                    if (splits[0] == "ADMIN")
                    {
                        ADMIN_START = int.Parse(splits[1]);
                    }
                }
            }
            catch (FormatException ex)
            {
                MessageBox.Show($"Error Parsing number in \"csvstart.txt\". Details: {ex.Message}");
            }

            configPath = Path.Combine(Environment.CurrentDirectory, "excelstart.txt");
            configText = File.ReadAllText(configPath);

            string[] _configPairs = configText.Split('\n');

            try
            {
                foreach (string configPair in _configPairs)
                {
                    string[] splits = configPair.Split('=');

                    if (splits[0] == "TITLEROW")
                    {
                        TITLEROW = int.Parse(splits[1]);
                    }
                }
            }
            catch (FormatException ex)
            {
                MessageBox.Show($"Error Parsing number in \"excelstart.txt\". Details: {ex.Message}");
            }

            if (ADMIN_START == -1)
            {
                MessageBox.Show("Error loading \"csvstart.txt\" config. ADMIN arguments incorrect");
                Application.Exit();
            }
            if (TITLEROW == -1)
            {
                MessageBox.Show("Error loading \"excelstart.txt\" config. TITLEROW arguments incorrect");
            }
        }
    }
}
