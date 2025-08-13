

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Fiverr_spreadsheet_project_1.Libraries
{
    internal class DataHandler
    {
        public class Job
        {
            public string workerName = string.Empty; 
            public string jobDescription = string.Empty;
            public double hours = 0.0;
        }

        public class Client
        {
            public string clientName = string.Empty;

            public List<Job> jobs = new List<Job>();

            public List<int> locations = new List<int>();

            public Dictionary<string, double> jobHours = new Dictionary<string, double>();

            public double totalJobHours = 0.0; // not including travel

            public double totalTravelHours = 0.0;
        }

        public void ConvertCsvToExcel(string csvPath, string excelPath, bool appending)
        {
            string[][] data = ReadCsv(csvPath);
            WriteToExcel(data, excelPath, appending);
        }

        public string[][] ReadCsv(string csvPath)
        {
            string[] lines;

            string[][] data = null;

            try
            {
                // Read all lines
                lines = File.ReadAllLines(csvPath);
            

                // Put into a 2D array in the format of [row][column]
                data = new string[lines.Length][];

                for (int i = 0; i < lines.Length; i++)
                {
                    // Split into columns
                    data[i] = lines[i].Split(',');
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error 003: {ex.Message}");
            }

            return data;
        }

        private double GetNumberOrThrow(IXLCell cell)
        {
            if (cell.IsEmpty())
                return 0.0;

            if (cell.IsMerged())
            {
                MessageBox.Show($"Cell {cell.Address} is merged. This may possibly cause an error.");
            }

            if (cell.DataType != XLDataType.Number)
                throw new Exception($"Error while totaling category hours: Expected a number in cell {cell.Address}, but found '{cell.GetValue<string>()}'");

            try
            {
                return cell.GetDouble(); // same as Value.GetNumber()
            }
            catch (InvalidCastException ex)
            {
                throw new Exception($"(Task cell) Invalid number format in cell {cell.Address}: {ex.Message}");
            }
        }


        public void WriteToExcel(string[][] data, string excelPath, bool appending)
        {
            bool errorHasOccurred = false;

            // Get information first before adding it
            // Make a list for all the clients
            List<Client> clients = new List<Client>();

            // For checking for multiple entries
            List<string> clientNames = new List<string>();

            // List of names
            List<string> workerNames = new List<string>();

            // To keep track of the hours each worker does for each client
            Dictionary<string, Dictionary<string, double>> workerHourInfo = new Dictionary<string, Dictionary<string, double>>();

            // !!! May need to change if he has data coming in with the same client on multiple rows
            // First we must find all the clients
            bool end = false;
            int i = 1;
            while (!end)
            {
                if (data.Length <= i)
                {
                    end = true;
                }
                else
                {
                    if (data[i].Length < 15)
                    {
                        i++;
                        continue;
                    }

                    string name = data[i][12];

                    if (clientNames.Contains(name))
                    {
                        clients[clientNames.IndexOf(name)].locations.Add(i);
                    }
                    else
                    {
                        Client newClient = new Client();

                        newClient.clientName = name;

                        newClient.locations.Add(i);

                        clients.Add(newClient);
                        clientNames.Add(name);
                    }
                }
                i++;
            }

            // Start at the first row and read in the first client
            // Then go through each client
            foreach (Client c in clients)
            {
                foreach (int loc in c.locations)
                {
                    // Get worker name, job name, and hours
                    Job newJob = new Job();

                    // Worker
                    string _workerName = newJob.workerName = data[loc][2];

                    if (!workerNames.Contains(_workerName))
                        workerNames.Add(_workerName);

                    // Job
                    // XXXXXXXXX Check indexes 15, 16, and 17 (to find which job) (DEPRECATED)
                    // Now using a config method
                    // Also use the new dictionary created from the config
                    string taskTitle = "";
                    string errorTracker = "ERROR 301";
                    try
                    {
                        if (data[loc][Loader.ADMIN_START] != string.Empty)
                        {
                            errorTracker = data[loc][Loader.ADMIN_START].Replace("\"", "");
                            taskTitle = newJob.jobDescription = Loader.configPairs[data[loc][Loader.ADMIN_START].Replace("\"", "")];
                        }
                        else if (data[loc][Loader.ADMIN_START + 1] != string.Empty)
                        {
                            errorTracker = data[loc][Loader.ADMIN_START + 1].Replace("\"", "");
                            taskTitle = newJob.jobDescription = Loader.configPairs[data[loc][Loader.ADMIN_START + 1].Replace("\"", "")];
                        }
                        else if (data[loc][Loader.ADMIN_START + 2] != string.Empty)
                        {
                            errorTracker = data[loc][Loader.ADMIN_START + 2].Replace("\"", "");
                            taskTitle = newJob.jobDescription = Loader.configPairs[data[loc][Loader.ADMIN_START + 2].Replace("\"", "")];
                        }
                        else
                        {
                            // If empty it goes in fieldwork
                            taskTitle = newJob.jobDescription = "Fieldwork";
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show($"\"{errorTracker}\" was not found when loading the config (config.txt) file."); // Removed details: " \n\nDetails: \"${e.Message}\" "
                        errorHasOccurred = true;
                    }

                    try
                    {
                        if (!workerHourInfo.ContainsKey(_workerName))
                        {
                            workerHourInfo[_workerName] = new Dictionary<string, double>();
                        }
                        if (!workerHourInfo[_workerName].ContainsKey(c.clientName))
                        {
                            workerHourInfo[_workerName][c.clientName] = 0;
                        }

                        // Hours
                        try
                        {
                            newJob.hours = double.Parse(data[loc][11]);
                        }
                        catch (InvalidCastException ex)
                        {
                            throw new Exception($"(Task cell in csv) Invalid number format in cell (Row: ${loc}, Col: L ): {ex.Message}");
                        }

                        if (newJob.jobDescription != "Travel") // Total anything that's not travel
                        {
                            c.totalJobHours += newJob.hours;
                        }
                        else // Total travel on its own (add both later)
                        {
                            c.totalTravelHours += newJob.hours;
                        }

                        workerHourInfo[_workerName][c.clientName] += newJob.hours;
                    } catch (Exception e)
                    {
                        errorHasOccurred = true;
                        MessageBox.Show($"Error Parsing job hours (Cell L{loc + 1}): {e.Message}");
                    }


                    c.jobs.Add(newJob);
                }

                // Now calculate the hours for each type of job
                foreach (Job job in c.jobs)
                {
                    if (!c.jobHours.ContainsKey(job.jobDescription))
                    {
                        c.jobHours[job.jobDescription] = 0;
                    }
                    c.jobHours[job.jobDescription] += job.hours;
                }
            }

            try
            {
                // Add values
                using (var workbook = new XLWorkbook(excelPath))
                {
                    var ws = workbook.Worksheet(1);

                    // Getting first empty row 
                    // Was using Column A now using B just incase number pattern isn't repeated
                    int lastRow = ws.Column(2).LastCellUsed().Address.RowNumber;

                    Dictionary<string, int> clientRowLocation = new Dictionary<string, int>();

                    // Make a dictionary of where each client is
                    for (int j = Loader.TITLEROW + 1; j < lastRow + 1; j++)
                    {
                        // We also must first find the start of the clients in the excel sheet
                        if (ws.Cell(j, 1).DataType != XLDataType.Number) // Meaning: if we don't see the number indicating the increment of clients
                        {
                            continue; // Keep going until we find it
                        }

                        var cell = ws.Cell(j, 2);
                        try
                        {
                            string clientName = cell.Value.GetText();

                            clientRowLocation[clientName] = j;
                        }
                        catch (InvalidCastException ex)
                        {
                            throw new Exception($"(Client name cell) Invalid number format in cell {cell.Address}: {ex.Message}");
                        }
                    }

                    int index = lastRow + 1;

                    // Get worker count (this is the amount of names displayed inside of the book)
                    List<string> workerNamesInBook = new List<string>();
                    int workerCount = 0;
                    if (!ws.Cell(2, 47).IsEmpty())
                    {
                        workerNamesInBook.Add(ws.Cell(2, 47).GetValue<string>());
                        workerCount++;

                        bool countFound = false;
                        while (!countFound)
                        {
                            if (!ws.Cell(2, 47 + (workerCount * 2)).IsEmpty())
                            {
                                workerNamesInBook.Add(ws.Cell(2, 47 + (workerCount * 2)).GetValue<string>());
                                workerCount++;
                            }
                            else
                            {
                                countFound = true;
                            }
                        }
                    }

                    // Now loop and add values
                    foreach (Client c in clients)
                    {
                        int ogIndex = -1;

                        if (clientRowLocation.ContainsKey(c.clientName))
                            ogIndex = clientRowLocation[c.clientName];

                        // First add the number
                        if (ogIndex == -1)
                            ws.Cell(index, 1).Value = index - 2;

                        // Next add the Client name
                        if (ogIndex == -1)
                            ws.Cell(index, 2).Value = c.clientName;

                        // Add to the location list
                        if (ogIndex == -1)
                        {
                            clientRowLocation.Add(c.clientName, index);
                        }


                        // Then job hours

                        // Planning
                        double planningTotal = GetNumberOrThrow(ws.Cell(ogIndex == -1 ? index : ogIndex, 8));
                        if (c.jobHours.ContainsKey("Planning"))
                        {
                            planningTotal += c.jobHours["Planning"];
                            ws.Cell(ogIndex == -1 ? index : ogIndex, 8).Value = planningTotal;
                        }

                        // Inventory Co.
                        double inventoryTotal = GetNumberOrThrow(ws.Cell(ogIndex == -1 ? index : ogIndex, 10));
                        if (c.jobHours.ContainsKey("Inventory Co."))
                        {
                            inventoryTotal += c.jobHours["Inventory Co."];
                            ws.Cell(ogIndex == -1 ? index : ogIndex, 10).Value = inventoryTotal;
                        }

                        // Interim
                        double interimTotal = GetNumberOrThrow(ws.Cell(ogIndex == -1 ? index : ogIndex, 12));
                        if (c.jobHours.ContainsKey("Inventory Co."))
                        {
                            interimTotal += c.jobHours["Inventory Co."];
                            ws.Cell(ogIndex == -1 ? index : ogIndex, 12).Value = interimTotal;
                        }

                        // Fieldwork
                        double fieldworkTotal = GetNumberOrThrow(ws.Cell(ogIndex == -1 ? index : ogIndex, 14));
                        if (c.jobHours.ContainsKey("Fieldwork"))
                        {
                            fieldworkTotal += c.jobHours["Fieldwork"];
                            ws.Cell(ogIndex == -1 ? index : ogIndex, 14).Value = fieldworkTotal;
                        }

                        // Reporting
                        double reportingTotal = GetNumberOrThrow(ws.Cell(ogIndex == -1 ? index : ogIndex, 16));
                        if (c.jobHours.ContainsKey("Reporting"))
                        {
                            reportingTotal += c.jobHours["Reporting"];
                            ws.Cell(ogIndex == -1 ? index : ogIndex, 16).Value = reportingTotal;
                        }

                        // Review and Sup
                        double revSupTotal = GetNumberOrThrow(ws.Cell(ogIndex == -1 ? index : ogIndex, 18));
                        if (c.jobHours.ContainsKey("Review and Sup"))
                        {
                            revSupTotal += c.jobHours["Review and Sup"];
                            ws.Cell(ogIndex == -1 ? index : ogIndex, 18).Value = revSupTotal;
                        }

                        // Meetings
                        double meetingsTotal = GetNumberOrThrow(ws.Cell(ogIndex == -1 ? index : ogIndex, 20));
                        if (c.jobHours.ContainsKey("Meetings"))
                        {
                            meetingsTotal += c.jobHours["Meetings"];
                            ws.Cell(ogIndex == -1 ? index : ogIndex, 20).Value = meetingsTotal;
                        }

                        // Processing
                        double processingTotal = GetNumberOrThrow(ws.Cell(ogIndex == -1 ? index : ogIndex, 22));
                        if (c.jobHours.ContainsKey("Processing"))
                        {
                            processingTotal += c.jobHours["Processing"];
                            ws.Cell(ogIndex == -1 ? index : ogIndex, 22).Value = processingTotal;
                        }

                        // Completion
                        double completionTotal = GetNumberOrThrow(ws.Cell(ogIndex == -1 ? index : ogIndex, 24));
                        if (c.jobHours.ContainsKey("Completion"))
                        {
                            completionTotal += c.jobHours["Completion"];
                            ws.Cell(ogIndex == -1 ? index : ogIndex, 24).Value = completionTotal;
                        }


                        // Total
                        double total = planningTotal + inventoryTotal + interimTotal + fieldworkTotal + reportingTotal + revSupTotal + meetingsTotal + processingTotal + completionTotal;
                        ws.Cell(ogIndex == -1 ? index : ogIndex, 26).Value = total;

                        // Travel
                        double travelTotal = 0.0;
                        var cell = ws.Cell(ogIndex == -1 ? index : ogIndex, 28);
                        if (!cell.IsEmpty())
                        {
                            if (cell.IsMerged())
                            {
                                MessageBox.Show($"Cell {cell.Address} is merged. This may possibly cause an error.");
                            }

                            if (cell.DataType != XLDataType.Number)
                                throw new Exception($"Error while totaling category hours: Expected a number in cell {cell.Address}, but found '{cell.GetValue<string>()}'");

                            try {
                                travelTotal = cell.Value.GetNumber() + c.totalTravelHours;
                            }
                            catch (InvalidCastException ex)
                            {
                                throw new Exception($"(Task cell) Invalid number format in cell {cell.Address}: {ex.Message}");
                            }
                        }
                        else
                            travelTotal = c.totalTravelHours;
                        ws.Cell(ogIndex == -1 ? index : ogIndex, 28).Value = travelTotal;

                        // Grand Total
                        double grandTotal = total + travelTotal;
                        ws.Cell(ogIndex == -1 ? index : ogIndex, 30).Value = grandTotal;


                        // Only need to increment if a client isn't already in the books
                        if (ogIndex == -1)
                            index++;
                    }

                    int formerWorkerCount = workerCount;
                    // Finally add the worker names to the list
                    foreach (string name in workerNames)
                    {
                        // Check if the worker is already in the book
                        if (workerNamesInBook.Contains(name))
                        {
                            continue; // Skip this worker if they are already in the book
                        }
                        else
                        {
                            // Add the name to workerNamesInBook (as they are just added)
                            // then the index of the name will act as a location finder
                            workerNamesInBook.Add(name);
                        }

                        // Shift every column past the last worker on to the right by N (maybe 2) columns
                        int N = 2;
                        ws.Column(47 + (workerCount * 2) + 1).InsertColumnsBefore(N);

                        // Format Columns
                        for (int j = 0; j < N; j++)
                        {
                            // int newIndex = (47 + (workerCount * 2) + 1) + j;
                            int newIndex = (47 + (workerCount * 2) - 1) + j;
                            ws.Column(newIndex).Style = ws.Column(47).Style;
                        }

                        // Add the workers name
                        int col = 47 + (workerCount * 2);
                        ws.Cell(Loader.TITLEROW, col).Value = name;
                        ws.Cell(Loader.TITLEROW, col).Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                        workerCount++;
                    }

                    foreach (var worker in workerHourInfo)
                    {
                        foreach (var client in worker.Value)
                        {
                            double temp = 0.0;
                            var cell = ws.Cell(clientRowLocation[client.Key], 47 + (workerNamesInBook.IndexOf(worker.Key) * 2));
                            if (!cell.IsEmpty())
                            {
                                if (cell.DataType != XLDataType.Number)
                                    throw new Exception($"Error while totaling worker hours: Expected a number in cell {cell.Address}, but found '{cell.GetValue<string>()}'");

                                try
                                {
                                    temp = cell.Value.GetNumber();
                                }
                                catch (InvalidCastException ex)
                                {
                                    throw new Exception($"(Worker cell) Invalid number format in cell {cell.Address}: {ex.Message}");
                                }
                            }
                            ws.Cell(clientRowLocation[client.Key], 47 + (workerNamesInBook.IndexOf(worker.Key) * 2)).Value = client.Value + temp;
                        }
                    }

                    if (!errorHasOccurred)
                    {
                        workbook.Save();
                    } else
                    {
                        return;
                    }
                }

                MessageBox.Show("Hour Report has successfully been added to the Analysis sheet");
            } 
            catch (ClosedXML.Excel.Exceptions.ClosedXMLException cx)
            {
                // ClosedXML error: corruption or file format issue
                MessageBox.Show($"Excel file appears corrupted or unreadable. Details: {cx.Message}");
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error 005: {ex.Message}");
                errorHasOccurred = true;
            }
        }
    }
}
