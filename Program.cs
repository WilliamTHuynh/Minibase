using System;
using System.Data.SqlClient;
using System.IO;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;

namespace MiniatureBase
{
    class Program
    {
        // Start by getting the data from the Google Sheet,  the Google Sheet being used is commented below.
        // https://docs.google.com/spreadsheets/d/192piBgNM8YBSnT5SOGHp8bmoE0kvDMFRvt-7HM7cQKQ/edit?usp=sharing

        // Using readonly strings because these shouldn't be modified Here we are looking specifically at the sheet FS1 + FS2 because we want to get all the names from that column. 
        // SpreadsheetID is found from the url of the spreadsheet. 
        static readonly string[] Scopes = {SheetsService.Scope.Spreadsheets};
        static readonly string applicationName = "Minibase";
        static readonly string spreadSheetID = "192piBgNM8YBSnT5SOGHp8bmoE0kvDMFRvt-7HM7cQKQ";
        static readonly string sheet = "FS1 + FS2";

        static SheetsService service;

        static void Main(string[] args)
        {

            // Asking the user for input to see how many rows would be read. 
            AskForInput();

            string userInput = ReadInput();

            // validMessage outputs any error messages if encountered. 
            string validMessage = "";
            if (IsValid(userInput, out validMessage) == false)
            {
                Console.WriteLine(validMessage);
                return; // exit function if invalid
            }
            //Start Reading and Moving data.

            // Now moving on to connect to the Azure Database and inputting data. 
            try
            {
                // Connect to the Azure server. This requires an Admins account and the server name, user input is required for this step.
                var cb = new SqlConnectionStringBuilder();


                cb.DataSource = "";
                cb.UserID = "";
                cb.Password = "";
                cb.InitialCatalog = "";

                // If connecting was successful then continue.; 
                using (var connection = new SqlConnection(cb.ConnectionString))
                {
                    connection.Open();
                    int num1, num2;

                    // Uses the user input from the beginning, which should be valid as it has already been checked by isValid. 
                    GetRange(userInput, out num1, out num2);
                    AddToMiniBase(connection,num1, num2);

                    // Once we are done using the server stop the connection. 
                    connection.Close();
                }
            }

            //If an error occurs write it to console. 
            catch (SqlException e) 
            {
                Console.WriteLine(e.ToString());
            }



            Console.WriteLine("View the report output here, then press any key to end the program...");
            Console.ReadKey();
            
        }

        static void AskForInput()
        {
            Console.WriteLine("Please enter number of rows up to 386 you want to transfer from Google Sheets to the Azure SQL Database." +
                "\nEither enter a single row or range of rows. Put the lower number first" +
                "\nFor example putting '2-10' will transfer rows 2 to 10 and '1' will transfer row 1.");
        }

        // Checks if input for row selection in valid. 
        /* Parameters: 
         *  input: a string that contains what the user has inputted. 
         *  msg: passed by reference, displays a error message if invalid. 
         * Output:
         *  If an error msg will be printed.
         * Returns: 
         *  A boolean indicating if the input is valid.
         */
        static bool IsValid(string input, out string msg)
        {
            // Starts valid_input as false and flags it as true whenever valid. This prevents the code from continuing by accident. 
            bool valid_input = false;
            string[] inputSplit;
            msg = "";

            // Checks if the input a single or two numbers. Here there are two numbers due to the -, and if it does continue splitting. 
            if (input.Contains("-")) 
            {

                // Splits the input at -
                inputSplit = input.Split('-');

                // Checks to see if there are two numbers. If so then check if the numbers go from smaller to bigger. 
                if (inputSplit.Length == 2)
                {
                    // If so then check if the numbers go from smaller to bigger.
                    if (Convert.ToInt32(inputSplit[0]) > Convert.ToInt32(inputSplit[1]))
                    {
                        valid_input = false;
                        msg = "Check your range";
                        return valid_input;
                    }

                    // check to see if both numbers don't exceed the amount of rows there are. 
                    for (int i = 0; i <= inputSplit.Length - 1; i++)
                    {
                        try
                        {
                            int.Parse(inputSplit[i].Trim());
                            if (Convert.ToInt32(inputSplit[i]) <= 386)
                            {
                                valid_input = true; 
                            }

                            else 
                            {
                                valid_input = false;
                                msg = "Row number exceed number of rows allowed.";
                                return valid_input;
                            }
                        }
                        catch
                        {
                            valid_input = false;
                            //valid_input remains false
                            msg = "Please check your input, numeric only.";
                            return valid_input;
                        }
                    }
                }

                else
                {
                    valid_input = false;
                    msg = "Please check your input.";
                    return valid_input;
                }
            }

            // Other case where there is only 1 number. 
            else 
            {
                input = input.Trim();
                try
                {
                    int.Parse(input);
                    valid_input = true;
                }
                catch
                {
                    valid_input = false;
                    //valid_input remains false
                    msg = "Please check your input, numeric only.";
                    return valid_input;
                }
            }
            return valid_input;
        }

        // Gets the range of rows to be checked over. If there is one number then num1 = num2
        /* Parameters: 
         *  input: a string that contains what the user has inputted. 
         *  num1: first number of input
         *  num2: second number of input
         * Output:
            num1, num2 
         * Returns: 
         *  None
         */
        static void GetRange(string input, out int num1, out int num2) {
            if (input.Contains("-"))
            {
                string[] inputSplit;
                // Splits the input by whitespace
                inputSplit = input.Split('-');
                num1 = Convert.ToInt32(inputSplit[0]);
                num2 = Convert.ToInt32(inputSplit[1]);
            }

            else
            {
                num1 = Convert.ToInt32(input);
                num2 = num1;
            }
        }

        // Recieves user input of the rows to be read. 
        /* Returns: 
         * String of user input
        */
        static string ReadInput()
        {
            // Asking the user for input to see how many rows would be read. 
            string input = Console.ReadLine();
            return input;
        }

        // Specifies rows to be added to the Miniature table of the Azure SQL database. 
        /* Parameters: 
         *  range1: an int that contains the first row to read. 
         *  range2: an int that contains the last row to read.
         * Output:
         *  Each row is added into the database and a successful attempt is printed to the console. 
         * Returns: 
         *  None
         */
        static void AddToMiniBase(SqlConnection connection,int range1, int range2)
        {
            // Create a GoogleCrendetial Object so we can use our key to connect to the API and manipulate the Google Sheet.
            GoogleCredential credential;

            // Read the google key (json file) to allow us to get credentials for read access. 
            using (var stream = new FileStream("boreal-freedom-312803-75784bd90c97.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);

            }

            // Takes the credential and application name 
            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });

            var range = $"{sheet}!{range1}:{range2}";
            var request = service.Spreadsheets.Values.Get(spreadSheetID, range);
            var response = request.Execute();
            var values = response.Values;

            // Declare the columns of the Miniature table so they can be inserted
            string miniatureName;
            string company;
            string material;
            int SKU;
            string sculpter;
            string genre;
            int scale;
            string DnDSize;

            // Checks if the reponse was null and if there are entries at all. If not, output each row and their their values. 
            if (values != null & values.Count > 0)
            {
                foreach (var row in values)
                {
                    // Each variable must be converted before being saved as they are seen as being a List Object
                    miniatureName = (string)row[0];
                    company = (string)row[1];
                    material = (string)row[2];
                    SKU = Convert.ToInt32((string)row[3]);
                    sculpter = (string)row[4];
                    genre = (string)row[5];
                    scale = Convert.ToInt32((string)row[6]);
                    DnDSize = (string)row[7];

                    // The addstring that is will be submitted to the Azure database.  Uses a format string to place all the values into the string. 
                    string addString =
                        String.Format(@"INSERT INTO Miniature (mName, mCompany, mMaterial, mSKU, mSculpter, mGenre, mScaleMM, mDnDSize, mIsPainted) 
                        VALUES ('{0}', '{1}', '{2}', {3}, '{4}', '{5}', '{6}', '{7}', 0);", miniatureName, company, material, SKU, sculpter, genre, scale, DnDSize);

                    Console.WriteLine(addString);

                    Submit_Tsql_NonQuery(connection, "1-Insert Miniature", addString);
                }
            }

            else
            {
                Console.WriteLine("No data found.");
            }
        }

        // Used from Microsoft's design your first database in C#.
        // https://docs.microsoft.com/en-us/azure/azure-sql/database/design-first-database-csharp-tutorial
        /* Parameters: 
         *  connection: represents a conenction to a SQL Server database. 
         *  tsqlPurpose: A string that to indicate what the function's objective is. 
         *  tsqlSourceCode: The SQL code that will be sent to the server to perform.
         *  parameterName: 
         *  parameterValue: 
         * Output:
         *  If the connection is succesful and there is no null then the miniature was successfully added to the database. 
         * Returns:
         *  None
         */
        static void Submit_Tsql_NonQuery(SqlConnection connection, string tsqlPurpose, string tsqlSourceCode,
            string parameterName = null,string parameterValue = null)
        {
            // Outputs the command that will be run to console. 
            Console.WriteLine();
            Console.WriteLine("=================================");
            Console.WriteLine("T-SQL to {0}...", tsqlPurpose);

            // Creates an Sql command that uses tsqlSourceCode and puts it to the specific connection. 
            using (var command = new SqlCommand(tsqlSourceCode, connection))
            {
                if (parameterName != null)
                {
                    command.Parameters.AddWithValue(
                        parameterName,
                        parameterValue);
                }
                int rowsAffected = command.ExecuteNonQuery();
                Console.WriteLine(rowsAffected + " = rows affected.");
            }
        }
    }
}

