using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;



namespace Shattered_Crystal_Hackathon
{
    class Program
    {
        static void Main(string[] args)
        {
            //string fullProgramFolderPath = Path.GetFullPath(@"CrystalReportProject/CrystalProject"); 
            string fullProgramFolderPath = @"D:\Users\A-Cron\Documents\Coding\CSharp\CrystalReportProject\CrystalProject";
            
            //string test = Path.GetFullPath("DatabaseConfig.txt");
            
            bool doesConfigFileExist = CheckIfTxtFileExists(fullProgramFolderPath);
            if(doesConfigFileExist == true)
            {
                Console.WriteLine("DatabaseConfig.txt file found");
            }
            else
            {    
                EndProgram("Error: Please add the DatabaseConfig.txt file to the project folder, then re-run this program");
            }

            string databaseConfigPath = @"D:\Users\A-Cron\Documents\Coding\CSharp\CrystalReportProject\CrystalProject\DatabaseConfig.txt";
            string[] lines = File.ReadAllLines(databaseConfigPath);
            DatabaseInfo newDatabaseInfo = new DatabaseInfo();
            if (lines.Length == 5)
            {  
                newDatabaseInfo.CrystalFilesFolder = Path.GetFullPath(lines[0]);
                newDatabaseInfo.ServerName = lines[1];
                newDatabaseInfo.DatabaseName = lines[2];
                newDatabaseInfo.UserId = lines[3];
                newDatabaseInfo.Password = lines[4];

                Console.WriteLine();
                Console.WriteLine("All fields found in DatabaseConfig.txt");
                Console.WriteLine("Crystal files folder path: " + newDatabaseInfo.CrystalFilesFolder);
                Console.WriteLine("Server name: " + newDatabaseInfo.ServerName);
                Console.WriteLine("Database name: " + newDatabaseInfo.DatabaseName);
                Console.WriteLine("UserID: " + newDatabaseInfo.UserId);
                Console.WriteLine("Password: " + newDatabaseInfo.Password);
            }
            else
            {
                EndProgram("Error: Please add all of the required database config information to the DatabaseConfig.txt file");
            }

            if (Directory.Exists(newDatabaseInfo.CrystalFilesFolder))
            {
                string[] filesInDirectory = Directory.GetFiles(newDatabaseInfo.CrystalFilesFolder);
                foreach(string file in filesInDirectory)
                {
                    ProcessCrystalReport(file, newDatabaseInfo.ServerName, newDatabaseInfo.DatabaseName, newDatabaseInfo.UserId, newDatabaseInfo.Password);
                }
            }
            else
            {
                EndProgram("Error: Crystal reports folder not found");
            }


            Console.WriteLine();
            Console.WriteLine("All reports updated");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            return;
            
        }

        public static bool CheckIfTxtFileExists(string path)
        {
            string[] filesInFolder = Directory.GetFiles(path);
            string fileName;
            foreach(string file in filesInFolder)
            {
                fileName = Path.GetFileName(file);
                if(fileName == "DatabaseConfig.txt")
                {
                    return true;
                }
            }
            return false;
        }

        public static void EndProgram( string message )
        {
            Console.WriteLine(message);
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        public static void ProcessCrystalReport(string filePath, string serverName, string databaseName, string userId, string password)
        {
            
            string fileExtension = Path.GetExtension(filePath);
            if(fileExtension == ".rpt")
            {
                ReportDocument boReportDocument = new ReportDocument();
                boReportDocument.Load(filePath);

                // Create a ConnectionInfo
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
            }
            else
            {
                return;
            }
        }

        public static void ModifyConnectionInfo(CrystalDecisions.CrystalReports.Engine.Database boDatabase, ConnectionInfo boConnectionInfo)
        {
            // Loop through each Table in the Database and apply the changes
            foreach (CrystalDecisions.CrystalReports.Engine.Table boTable in boDatabase.Tables)
            {
                TableLogOnInfo boTableLogOnInfo = (TableLogOnInfo)boTable.LogOnInfo.Clone();
                boTableLogOnInfo.ConnectionInfo = boConnectionInfo;

                boTable.ApplyLogOnInfo(boTableLogOnInfo);

                // The location may need to be updated if the fully qualified name changes.
                //boTable.Location = "";    
            }
        }
    }
}
