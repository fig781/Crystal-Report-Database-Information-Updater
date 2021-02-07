using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace Shattered_Crystal_Hackathon
{
    class Program
    {
        static void Main(string[] args)
        {
            //Find DatabaseConfig.txt file
            bool proceed = false;
            string txtFilePathInput = "";
            while(proceed == false){
                Console.WriteLine("Enter the path of the DatabaseConfig.txt file below. Ex: C:\\Users\\Bob\\Documents\\DatabaseConfig.txt: ");
                txtFilePathInput = Console.ReadLine();

                if (File.Exists(txtFilePathInput))
                {
                    Console.WriteLine("DatabaseConfig.txt file found.");
                    proceed = true;
                }
                else
                {
                    Console.WriteLine("Error: DatabaseConfig.txt not found. Make sure the file path is correct");
                    Console.WriteLine();
                }
            }

            //Read DatabaseConfig.txt file 
            string[] lines = File.ReadAllLines(txtFilePathInput);
            DatabaseInfo newDatabaseInfo = new DatabaseInfo();
            if (lines.Length >= 5)
            {  
                newDatabaseInfo.CrystalFilesFolder = lines[0];
                newDatabaseInfo.ServerName = lines[1];
                newDatabaseInfo.DatabaseName = lines[2];
                newDatabaseInfo.UserId = lines[3];
                newDatabaseInfo.Password = lines[4];

                Console.WriteLine();
                Console.WriteLine("All fields found in DatabaseConfig.txt");
                Console.WriteLine();
                Console.WriteLine("Crystal files folder path: " + newDatabaseInfo.CrystalFilesFolder);
                Console.WriteLine("Server name: " + newDatabaseInfo.ServerName);
                Console.WriteLine("Database name: " + newDatabaseInfo.DatabaseName);
                Console.WriteLine("UserID: " + newDatabaseInfo.UserId);
                Console.WriteLine("Password: " + newDatabaseInfo.Password);
                Console.WriteLine();
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("Error: Add all of the required database config information to the DatabaseConfig.txt file");
                Console.WriteLine("The DatabaseConfig.txt file should look something like this:");
                Console.WriteLine("Crystal Reports Folder Path");
                Console.WriteLine("Server Name");
                Console.WriteLine("Database Name");
                Console.WriteLine("User Id");
                EndProgram("Password");
            }

            //Proceed or end the program
            Console.Write("Proceed (Y/N): ");
            string input = Console.ReadLine().ToUpper();
            if(input != "Y" && input !="YES")
            {
                Console.WriteLine();
                EndProgram("Ending program");
            }

            //Check for and if need be, make a Log folder
            string programFolderPath = txtFilePathInput.Remove(txtFilePathInput.Length - 18);
            string logFolderPath = programFolderPath + "Logs";
            if (!Directory.Exists(logFolderPath))
            {
                Directory.CreateDirectory(logFolderPath);
            }

            //Process and alter crystal files
            List<string> logData = new List<string>();
            if (Directory.Exists(newDatabaseInfo.CrystalFilesFolder))
            {
                Console.WriteLine();
                Console.WriteLine("Altering crystal files:");
                int rptFileCount = 0;
                string[] filesInDirectory = Directory.GetFiles(newDatabaseInfo.CrystalFilesFolder);

                int filesProcessed = 1;
                foreach (string file in filesInDirectory)
                {
                    string updateLog = ProcessCrystalReport(filesProcessed, file, newDatabaseInfo.ServerName, newDatabaseInfo.DatabaseName, newDatabaseInfo.UserId, newDatabaseInfo.Password);
                    logData.Add(updateLog);
                    rptFileCount++;
                    filesProcessed++;
                }
                if(rptFileCount == 0)
                {
                    EndProgram("Error: No .rpt files found. Add .rpt files to this directory and re-run the program");
                }
            }
            else
            {
                EndProgram("Error: Crystal reports folder not found. The folder path may be written incorrectly in the DatabaseConfig.txt file");
            }

            //Generate log file           
            string logFileName = "Log-CrystalDatabaseUpdater-" + DateTime.Now.ToString("MM.dd.yyyy-hh.mm.sstt") + ".txt";
            string logFilePath = Path.Combine(logFolderPath, logFileName);
            using(FileStream fs = File.Create(logFilePath))
            {
                AddText(fs, "Variables");
                AddText(fs, "\r\nServer Name: " + newDatabaseInfo.ServerName);
                AddText(fs, "\r\nDatabase Name: " + newDatabaseInfo.DatabaseName);
                AddText(fs, "\r\nUser ID: " + newDatabaseInfo.UserId);
                AddText(fs, "\r\nPassword: " + newDatabaseInfo.Password);
                AddText(fs, "\r\n\r\n");

                for(int i=0; i < logData.Count; i++)
                {
                    AddText(fs, logData[i] + "\r\n");
                }
            } 

            Console.WriteLine();
            EndProgram("Program finished");
        }

        public static void AddText(FileStream fs, string value)
        {
            byte[] info = new UTF8Encoding(true).GetBytes(value);
            fs.Write(info, 0, info.Length);
        }

        public static void EndProgram( string message )
        {
            Console.WriteLine(message);
            Console.Write("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        public static string ProcessCrystalReport(int timesRan, string filePath, string serverName, string databaseName, string userId, string password)
        {
            string fileName ="Unknown";
            string returnVariables = "";
            timesRan.ToString();

            try
            {
                string fileExtension = Path.GetExtension(filePath);
                if (fileExtension == ".rpt")
                {
                    using(ReportDocument boReportDocument = new ReportDocument())
                    {
                        
                        boReportDocument.Load(filePath);

                        ConnectionInfo boConnectionInfo = new ConnectionInfo();
                        boConnectionInfo.ServerName = serverName;
                        boConnectionInfo.DatabaseName = databaseName;
                        boConnectionInfo.UserID = userId;
                        boConnectionInfo.Password = password;

                        ModifyConnectionInfo(boReportDocument.Database, boConnectionInfo);

                        foreach (ReportDocument boSubreport in boReportDocument.Subreports)
                        {
                            ModifyConnectionInfo(boSubreport.Database, boConnectionInfo);
                        }

                        fileName = Path.GetFileName(filePath);
                        Console.WriteLine("{0} - Updated", fileName);

                        //Write each section of the log file
                        string[] connectionVariables = GetConnectionInfo(boReportDocument.Database);
                        returnVariables = "\r\n" + "Server Name: " + connectionVariables[0] + "\r\n" + "Database Name: " + connectionVariables[1] + "\r\n" + "User ID: " + connectionVariables[2] + "\r\n" + "Password: " + boConnectionInfo.Password + "\r\n";
                    }

                    return timesRan + ". " + DateTime.Now.ToString() + " - " + fileName + " - Updated:" + returnVariables;
                }
                else
                {
                    return DateTime.Now.ToString() + " - " + fileName + " - Skipped";
                }
            }
            catch(Exception e)
            {
                fileName = Path.GetFileName(filePath);
                Console.WriteLine("{0} - Error encountered", fileName);
                Console.WriteLine(e);
                Console.WriteLine();
                return timesRan + ". " + DateTime.Now.ToString() + " - " + fileName + " - Error encountered - " + e;
            }
        }

        public static void ModifyConnectionInfo(Database boDatabase, ConnectionInfo boConnectionInfo)
        {
            // Loop through each Table in the Database and apply the changes
            foreach (Table boTable in boDatabase.Tables)
            {              
                TableLogOnInfo boTableLogOnInfo = (TableLogOnInfo)boTable.LogOnInfo.Clone();
                boTableLogOnInfo.ConnectionInfo = boConnectionInfo;
                boTable.ApplyLogOnInfo(boTableLogOnInfo);

                // The location may need to be updated if the fully qualified name changes.
                //boTable.Location = "";    
            }
        }

        public static string[] GetConnectionInfo(Database boDatabase)
        {
            //Loops through once and grabs the table info
            string[] connectionVariables = new string[4];
            foreach (Table boTable in boDatabase.Tables)
            {
                TableLogOnInfo boTableLogOnInfo = (TableLogOnInfo)boTable.LogOnInfo.Clone();
                connectionVariables[0] = boTableLogOnInfo.ConnectionInfo.ServerName;
                connectionVariables[1] = boTableLogOnInfo.ConnectionInfo.DatabaseName;
                connectionVariables[2] = boTableLogOnInfo.ConnectionInfo.UserID;
                connectionVariables[3] = boTableLogOnInfo.ConnectionInfo.Password;
                break;

                //Known bug, Password cannot be retrieved. It will be an empy string. I know it is being set though. Cannot figure out fix.
            }

            return connectionVariables;
        }
    }
}
