using System;
using System.IO;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace Shattered_Crystal_Hackathon
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter the path of the DatabaseConfig.txt file below. Ex: C:\\Users\\Bob\\Documents\\DatabaseConfig.txt: ");
            string txtFilePathInput = Console.ReadLine();
            Console.WriteLine();

            if (File.Exists(txtFilePathInput))
            {
                Console.WriteLine("DatabaseConfig.txt file found.");               
            }
            else
            {
                EndProgram("Error: DatabaseConfig.txt not found. Make sure the file path is correct");
            }

            //string databaseConfigPath = Path.GetFullPath("DatabaseConfig.txt");
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


            Console.Write("Proceed (Y/N): ");
            string input = Console.ReadLine().ToUpper();
            if(input != "Y")
            {
                Console.WriteLine();
                EndProgram("Ending program");
            }

            if (Directory.Exists(newDatabaseInfo.CrystalFilesFolder))
            {
                Console.WriteLine();
                Console.WriteLine("Altering crystal files:");
                int rptFileCount = 0;
                string[] filesInDirectory = Directory.GetFiles(newDatabaseInfo.CrystalFilesFolder);
                foreach(string file in filesInDirectory)
                {
                    ProcessCrystalReport(file, newDatabaseInfo.ServerName, newDatabaseInfo.DatabaseName, newDatabaseInfo.UserId, newDatabaseInfo.Password);
                    rptFileCount++;
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

            Console.WriteLine();
            EndProgram("Program finished");
        }

        public static void EndProgram( string message )
        {
            Console.WriteLine(message);
            Console.Write("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        public static int ProcessCrystalReport(string filePath, string serverName, string databaseName, string userId, string password)
        {
            try
            {
                string fileExtension = Path.GetExtension(filePath);
                if (fileExtension == ".rpt")
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

                    Console.WriteLine("{0} - Updated", Path.GetFileName(filePath));
                    return 1;
                }
                else
                {
                    return 0;
                }
            }
            catch(Exception e)
            {
                Console.WriteLine("{0} - Error encountered", Path.GetFileName(filePath));
                Console.WriteLine(e);
                Console.WriteLine();
                return 0;
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
