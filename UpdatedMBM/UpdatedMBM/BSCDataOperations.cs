using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using MinimalisticTelnet;
using OfficeOpenXml;
using System.Drawing;
using System.Xml;
using System.IO;
using OfficeOpenXml.Drawing;
//using Excel = Microsoft.Office.Interop.Excel;

namespace UpdatedMBM
{
    public class BSCDataOperations
    {
        public Dictionary<string, Dictionary<string, string>> BSCLogInInfo = new Dictionary<string, Dictionary<string, string>>();
        public Dictionary<string,List<Dictionary<string,string>>> CommandData = new Dictionary<string,List< Dictionary<string, string>>>(); 
        List<string> interruptStrings = new List<string>(); 
        List<string> commandEndText = new List<string>(); 
        public List<string> CommandLog = new List<string>();
        List<Dictionary<string,string>> ReportData = new List<Dictionary<string, string>>(); 
        public BSCDataOperations()
        {
            LoadBSCData();
            LoadInterruptStrings();
            LoadCommandEndText();
        }

        private void LoadInterruptStrings()
        {
            using (StreamReader streamReader = new StreamReader("InterruptText.txt"))
            {
                string temp = string.Empty;
                while ((temp = streamReader.ReadLine()) != null)
                {
                    interruptStrings.Add(temp.Trim());
                }
            }
        }

        private void LoadBSCData()
        {
            using (StreamReader aReader = new StreamReader("ip_map.csv"))
            {
                string temp = string.Empty;
                while ((temp = aReader.ReadLine()) != null)
                {
                    string[] data = temp.Split(',');
                    Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                    aDictionary.Add("CNUM", data[0]);
                    aDictionary.Add("BSCName", data[1]);
                    aDictionary.Add("IP", data[2]);
                    aDictionary.Add("User", data[3]);
                    aDictionary.Add("Password", data[4]);

                    if (!BSCLogInInfo.ContainsKey(data[0]))
                    {
                        BSCLogInInfo.Add(data[0],aDictionary);
                    }


                }
            }
        }

        public void LoadInputFile(string textFileName)
        {
            using (StreamReader streamReader = new StreamReader(textFileName))
            {
                string temp = string.Empty;
                while ((temp = streamReader.ReadLine()) != null)
                {
                    string[] data = temp.Split('@');
                    if (data.Length > 1)
                    {
                        if (CommandData.ContainsKey(data[1]))
                        {
                            Dictionary<string, string> aCommand = new Dictionary<string, string>();
                            aCommand.Add("Command", data[0]);
                            CommandData[data[1]].Add(aCommand);
                        }
                        else
                        {
                            List<Dictionary<string, string>> aList = new List<Dictionary<string, string>>();
                            Dictionary<string, string> aCommand = new Dictionary<string, string>();
                            aCommand.Add("Command", data[0]);
                            aList.Add(aCommand);
                            CommandData.Add(data[1], aList);
                        }
                    }

                }

                streamReader.Close();
            }
            
        }
        public void LoadInputFile(string excelFileName,string commandSheetName, string commandColumnName, string bscColumnName)
        {

        }

        public void RunScript()
        {
            
        }

        public void ExecuteCommandInBSC()
        {
            //if (File.Exists("output.csv"))
            //{
            //    File.Delete("output.csv");
            //}
            //StreamWriter sw = new StreamWriter("output.csv");
            //sw.Close();
            if (File.Exists("log.log"))
            {
                File.Delete("log.log");
            }
            StreamWriter sw = new StreamWriter("log.log");
            sw.Close();

            int index = 1;
            foreach (KeyValuePair<string, List<Dictionary<string, string>>> aBSCCommandPair in CommandData)
            {

                if (BSCLogInInfo.ContainsKey(aBSCCommandPair.Key))
                {
                    try
                    {
                        TelnetConnection tc = new TelnetConnection(BSCLogInInfo[aBSCCommandPair.Key]["IP"], 23);
                        bool LoggInSuccessFull = false;
                        string s = tc.Login(BSCLogInInfo[aBSCCommandPair.Key]["User"],
                            BSCLogInInfo[aBSCCommandPair.Key]["Password"], 10, out LoggInSuccessFull);
                        Console.Write(s);
                        CommandLog.Add(s);
                        string prompt = s.TrimEnd();
                        prompt = "\r";
                        foreach (Dictionary<string, string> dictionary in aBSCCommandPair.Value)
                        {
                            try
                            {
                                if (tc.IsConnected)
                                {
                                  
                                    prompt = dictionary["Command"];
                                    Console.WriteLine("Sending Command("+index+"): " + prompt);
                                    index++;
                                    string msg = tc.ExecuteCommand(prompt);
                                    Console.Write(msg);
                                    CommandLog.Add(msg);
                                    string finalComment = GetCommandFinalComment(msg);
                                    string errors = GetErrors(msg);
                                    dictionary.Add("LogData", msg);
                                    dictionary.Add("FinalComment", finalComment);
                                    dictionary.Add("Errors", errors);
                                    dictionary.Add("BSC", aBSCCommandPair.Key);
                                   

                                    if (ShouldInterrupt(msg))
                                    {
                                        string inte = tc.Interrupt();
                                        Console.WriteLine(inte);
                                        CommandLog.Add(inte);
                                        finalComment = GetCommandFinalComment(inte);
                                        dictionary["FinalComment"] = finalComment;
                                    }
                                    ReportData.Add(dictionary);
                                    //Thread.Sleep(10);

                                    if (ReportData.Count > 100)
                                    {
                                        //AppendDataInCSVFile(ReportData);
                                        //ReportData.Clear();
                                        WriteCommadLog();
                                    }
                                }
                            }
                            catch (Exception exception)
                            {
                                #region update dictionary
                                if (dictionary.ContainsKey("LogData"))
                                {
                                    dictionary["LogData"] = "";
                                }
                                else
                                {
                                    dictionary.Add("LogData", "");
                                }

                                if (dictionary.ContainsKey("FinalComment"))
                                {
                                    dictionary["FinalComment"] = "Exception Occured, Command Not Executed. Try it manually";
                                }
                                else
                                {
                                    dictionary.Add("FinalComment", "Exception Occured, Command Not Executed. Try it manually");
                                }

                                if (dictionary.ContainsKey("Errors"))
                                {
                                    dictionary["Errors"] = exception.Message;
                                }
                                else
                                {
                                    dictionary.Add("Errors", exception.Message);
                                }
                                if (dictionary.ContainsKey("BSC"))
                                {
                                    dictionary["BSC"] = aBSCCommandPair.Key;
                                }
                                else
                                {
                                    dictionary.Add("BSC", aBSCCommandPair.Key);
                                }

                                #endregion

                                ReportData.Add(dictionary);
                            }
                        }

                    }
                    catch (Exception exception)
                    {
                        
                        foreach (Dictionary<string, string> dictionary in aBSCCommandPair.Value)
                        {
                            dictionary.Add("LogData", "");
                            dictionary.Add("FinalComment", "Unabale to Connect BSC");
                            dictionary.Add("Errors", exception.Message);
                            dictionary.Add("BSC", aBSCCommandPair.Key);
                            ReportData.Add(dictionary);

                        }
                    }

                }
                else
                {
                    Console.WriteLine("BSC: " + aBSCCommandPair.Key + " login data not found.");
                    CommandLog.Add("BSC: " + aBSCCommandPair.Key + " login data not found.");
                }

            }

            WriteCommadLog();
            AppendDataInExelFile(ReportData, "Output");

            //AppendDataInCSVFile(ReportData);
            ReportData.Clear();
            
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("OPERATION COMPLETE...");
            
        }

        private string GetTimerData(string paramName,string msg)
        {
            string timer = string.Empty;

            string[] splittedMSG = msg.Split('\n');
            foreach (string aSplittedMsg in splittedMSG)
            {
                if (aSplittedMsg.Contains(paramName))
                {
                    string data = aSplittedMsg.Trim().Substring(aSplittedMsg.Trim().LastIndexOf(". ") +1);
                    return data;
                }
            }


            return timer;
        }
        private string GetErrors(string msg)
        {
            string error = string.Empty;
            string[] lineMsg = Regex.Split(msg, "\n");
            foreach (string aLineMsg in lineMsg)
            {
                if (aLineMsg.Contains("/*") && aLineMsg.Contains("*/"))
                {
                    error += aLineMsg + "\n";
                }
            }


            return error;
        }

        private void LoadCommandEndText()
        {
            using (StreamReader aReader = new StreamReader("commandendtext.txt"))
            {
                string temp = string.Empty;
                while ((temp=aReader.ReadLine())!= null)
                {
                    commandEndText.Add(temp);
                }

                aReader.Close();
            }
        }
        private string GetCommandFinalComment(string msg)
        {
            string[] data = msg.Split('\n');
            foreach (string aData in data)
            {
                foreach (string aEndText in commandEndText)
                {
                    if (aData.ToLower().Contains(aEndText.ToLower()))
                    {
                        return aData;
                    }
                }
            }
            return "";
        }

        

        private void WriteCommadLog()
        {
            
            using (StreamWriter sw = new StreamWriter("log.log", true))
            {
                foreach (string aLog in CommandLog)
                {
                    sw.WriteLine(aLog);
                }
                sw.Close();
            }
            CommandLog.Clear();
        }

        
        public bool ShouldInterrupt(string msg)
        {
            //using command

           

            string tMsg = msg.Trim();
            List<string> aList = Regex.Split(tMsg,"\r\n").ToList();
            string firstData = aList.First();
            string lastData = aList.Last();
           
            if (lastData.Contains("\b"))
            {
                string temp = lastData.Substring(0, lastData.IndexOf("\b"));
                string anoTemp = string.Empty;

                int index = lastData.IndexOf("\b") + 2;
                
                if (index < lastData.Length)
                {
                    anoTemp = lastData.Substring(lastData.IndexOf("\b") + 2);
                    lastData = temp + anoTemp;

                }
                else
                {
                    lastData = aList.Last();
                }
                
            }

            if (firstData.Trim().Contains(lastData.Trim().Substring(0, lastData.Trim().Length - 1)))
            {
                return true;
            }

           
            foreach (string interruptString in interruptStrings)
            {
                if (msg.ToLower().Contains(interruptString.ToLower()))
                {
                    return true;
                }
            }


            if (msg.ToLower().Contains("command ex") && !lastData.Contains(';'))
            {
                return false;
            }
            return false;
        }

        public void AppendDataInCSVFile(List<Dictionary<string, string>> data)
        {
           // string completeFileName = string.Empty;

            
            string outFile = "output.csv";
           
            using (StreamWriter sw = new StreamWriter(outFile,true))
            {
                string lineData = string.Empty;
                foreach (KeyValuePair<string, string> keyValuePair in data[0])
                {
                    lineData += keyValuePair.Key + ",";
                    
                }
                lineData = lineData.TrimEnd(',');
                sw.WriteLine(lineData);
                sw.Close();
            }

            foreach (Dictionary<string, string> dictionary in data)
            {
                string lineData = string.Empty;
                foreach (KeyValuePair<string, string> keyValuePair in dictionary)
                {
                    lineData += keyValuePair.Value + ",";
                }
                lineData = lineData.TrimEnd(',');
                using (StreamWriter sw = new StreamWriter(outFile, true))
                {
                    sw.WriteLine(lineData);
                    sw.Close();
                }
            }

        }

        public string AppendDataInExelFile(List<Dictionary<string, string>> data, string sheetName)
        {
            string completeFileName = string.Empty;

            int prefix = 0;
            string outFile = "output.xls";
            string tempFileName = outFile;
            

            completeFileName = Path.Combine(tempFileName);

            FileInfo newFile = new FileInfo(completeFileName);
            
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(completeFileName);
            }



            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
                int excelRowIndex = 1;
                int excelColumnIndex = 1;
                
                foreach (KeyValuePair<string, string> keyValuePair in data[0])
                {
                    worksheet.Cells[excelRowIndex, excelColumnIndex].Value = keyValuePair.Key.ToUpper();
                    excelColumnIndex++;
                }


                foreach (Dictionary<string, string> dictionary in data)
                {
                    excelRowIndex++;
                    excelColumnIndex = 1;
                    foreach (KeyValuePair<string, string> keyValuePair in dictionary)
                    {
                        int intData;
                        bool result = int.TryParse(keyValuePair.Value, out intData);
                        if (result)
                        {
                            worksheet.Cells[excelRowIndex, excelColumnIndex].Value = intData;
                        }
                        else
                        {
                            string dd = XmlConvert.EncodeName(keyValuePair.Value);
                            worksheet.Cells[excelRowIndex, excelColumnIndex].Value = dd;
                           
                        }
                        excelColumnIndex++;
                    }
                }


                if (data.Count > 0)
                {
                    worksheet.Column(1).Width = 25;
                    worksheet.Column(2).Width = 50;
                    worksheet.Column(3).Width = 40;
                    worksheet.Column(4).Width = 50;

                    using (var range = worksheet.Cells[1, 1, 1, data[0].Count])
                    {
                        //Setting bold font
                        range.AutoFilter = true;
                        range.Style.Font.Bold = true;

                        //Setting fill type solid
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        //Setting background color dark blue
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.BurlyWood);
                        //Setting font color
                        range.Style.Font.Color.SetColor(Color.White);
                        
                    }
                    using (var range = worksheet.Cells[1, 1, data.Count + 1, data[0].Count])
                    {
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                }
                package.Workbook.Properties.Author = "All Jubayer Mohammad Mahamudunnabi";
                package.Save();
            }
            return completeFileName;
        }
    }
}
