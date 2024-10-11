/*
Before You Begin
Instructions for Setting Up Google Sheets API with OAuth 2.0:
1. Go to the Google Cloud Console:
   - Open your web browser and navigate to the Google Cloud Console: https://console.cloud.google.com/
2. Create a New Project:
   - Click on the project dropdown at the top of the page and select "New Project."
   - Enter a name for your project and click "Create."
3. Enable the Google Sheets API:
   - In the Google Cloud Console, go to the "APIs & Services" > "Library."
   - Search for "Google Sheets API" and click on it.
   - Click the "Enable" button.
4. Set Up the OAuth Consent Screen:
   - Go to "APIs & Services" > "OAuth consent screen."
   - Choose "External" and click "Create."
   - Fill in the required fields (App name, User support email, Developer contact information).
   - Click "Save and Continue."
   - On the "Scopes" screen, click "Add or Remove Scopes."
     - In the "Add scopes" window, search for and select the following scope:
       - https://www.googleapis.com/auth/spreadsheets
     - Check the box beside "See, edit, create and delete all your Google Sheets spreadsheets".
     - Click "Update" to add the selected scope.
   - Click "Save and Continue" through the remaining steps.
   - On the "Test users" screen, add the email addresses of the users you want to allow access during the testing phase.
   - Click "Save and Continue" and then "Back to Dashboard."
5. Create OAuth 2.0 Credentials:
   - Go to "APIs & Services" > "Credentials."
   - Click on "Create Credentials" and select "OAuth Client ID."
   - Choose "Desktop app" as the application type and click "Create."
   - Download the JSON file containing your credentials and save it as credentials.json in your project directory, in the same directory as your Program.cs file.
6. Create a Google Sheet:
    - Open Google Sheets and create a new spreadsheet.
    - At the bottom of the Google Sheets window, rename the sheet from "Sheet1" to a name of your choice (Pets in this example, ideally use the plural version of the name of your class below).
    - Add a heading row to your sheet and ensure the headings match the properties of the class below (the Pet class in this example).
    - Note down the ID of the Google Sheet (found in the URL after "/d/" and before /edit).
7. Install the Google Sheets API Client Library:
   - Open your project in Visual Studio.
   - Go to the "Tools" menu at the top of the screen.
   - Select "NuGet Package Manager" > "Package Manager Console."
   - In the "Package Manager Console" window, type the following command and press Enter:
     Install-Package Google.Apis.Sheets.v4
   - Wait for the installation to complete. You should see messages indicating that the package has been installed successfully.
8. Set Up Your Code:
   - Use the downloaded credentials.json file and saves it into your project directory.
   - IMPORTANT: It is important that you do NOT commit the credentials.json file to source control, as it contains sensitive information.
   - Add the following line to your .gitignore file to prevent the credentials file from being committed to source control:
     credentials.json
   - Follow the provided code example to initialize the Google Sheets API client and interact with your Google Sheet.
*/

using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System.Reflection;
using System.Collections;
using System.Linq.Expressions;
using static Program;

internal class Program
{
    public class Pet
    {
        // Define the primary key for the class
        [GoogleSheet<Pet>.PrimaryKey]
        public string Name { get; set; }

        public string Species { get; set; }
        public int Age { get; set; }
        public double Price { get; set; }
        public string Color { get; set; }
    }

    static async Task Main(string[] args)
    {
        // Get the path to the current directory and the project directory
        string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string projectDirectory = Directory.GetParent(currentDirectory).Parent.Parent.Parent.FullName;

        //Check if the credentials file exists in the project directory
        string credentialsPath = Path.Combine(projectDirectory, "credentials.json");

        if (File.Exists(credentialsPath))
        {
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine("WARNING: Protect your sensitive information!");
            Console.WriteLine("\nThe 'credentials.json' file is inside your project directory, which is why you are seeing this message.");
            Console.WriteLine("\nThis 'credentials.json' file contains sensitive data and must NOT be committed to source control.");
            Console.WriteLine("\nPlease ensure that you add the file to your '.gitignore'.");
            Console.WriteLine("\nFor more details, refer to the instructions at the top of this example or speak with your professor.");
            Console.WriteLine("\nIMPORTANT: Do NOT commit 'credentials.json' to avoid exposing sensitive information.");
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine("\nPress any key to continue...");
            Console.ReadKey();
        }

        ////Create a GoogleSheet instance for the Pet class
        //var sheet = new GoogleSheet<Pet>("application_name_here", "google_sheet_id_here", "google_sheet_name_here");

        //Example:
        //var sheet = new GoogleSheet<Pet>("Pet Shop", "google_sheet_id_here", "Pets"); // Replace "google_sheet_id_here" with your Google Sheet ID (found in the URL after "/d/" and before /edit).


        ////IMPORTANT: Ensure that your Google Sheet is blank BEFORE calling the LoadAsync method for the first time.

        ////Load data from the Google Sheet
        //await sheet.LoadAsync();

        //// Check if a pet named "Fluffy" exists, and add it if it doesn't
        //if (!sheet.Any(pet => pet.Name == "Fluffy"))
        //{
        //    var newPet = new Pet
        //    {
        //        Name = "Fluffy",
        //        Species = "Cat",
        //        Age = 3,
        //        Price = 100,
        //        Color = "White"
        //    };

        //    sheet.Add(newPet);
        //}

        //// Check if a pet named "Buddy" exists, and add it if it doesn't
        //if (!sheet.Any(pet => pet.Name == "Buddy"))
        //{
        //    var newPet = new Pet
        //    {
        //        Name = "Buddy",
        //        Species = "Dog",
        //        Age = 1,
        //        Price = 300,
        //        Color = "Brown"
        //    };

        //    sheet.Add(newPet);
        //}

        //// Update the age of the pet named "Buddy"
        //var buddy = sheet.FirstOrDefault(pet => pet.Name == "Buddy");
        //buddy.Age += 1;

        //// Save changes to the Google Sheet
        //await sheet.SaveAsync();

        ////List all pets in the Google Sheet
        //Console.WriteLine("------------------------------------------------");
        //Console.WriteLine("Name\tSpecies\tAge\tPrice\tColor");
        //Console.WriteLine("------------------------------------------------");
        //foreach (Pet pet in sheet)
        //{
        //    Console.WriteLine($"{pet.Name}\t{pet.Species}\t{pet.Age}\t{pet.Price}\t{pet.Color}");
        //}
        //Console.WriteLine("------------------------------------------------");
    }
}

public class GoogleSheet<T> : IEnumerable<T> where T : class, new()
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class PrimaryKeyAttribute : Attribute
    {
    }

    private readonly string applicationName;
    private readonly string spreadsheetId;
    private readonly string sheetName;
    private SheetsService sheetsService;
    private Dictionary<string, T> changeTracker = new Dictionary<string, T>();
    private List<T> sheetData = new List<T>();
    private Dictionary<string, int> rowIndexByPrimaryKey = new Dictionary<string, int>();

    // Event triggered when an item is added
    public event Action<T> ItemAdded;

    public GoogleSheet(string applicationName, string spreadsheetId, string sheetName)
    {
        this.applicationName = applicationName;
        this.spreadsheetId = spreadsheetId;
        this.sheetName = sheetName;

        InitializeServiceAsync().ContinueWith(async task =>
        {
            await AddHeadingsAsync(); // Check and add headings if necessary
        });

        // Subscribe to the ItemAdded event and automatically handle it
        ItemAdded += OnItemAdded;
    }

    private async Task InitializeServiceAsync()
    {
        string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string projectDirectory = Directory.GetParent(currentDirectory).Parent.Parent.Parent.FullName;

        string credentialsPath = "credentials.json";

        if (!File.Exists(credentialsPath))
        {
            credentialsPath = Path.Combine(projectDirectory, "credentials.json");
        }
        else
        {
            throw new Exception("Credentials file not found.");
        }

        UserCredential userCredential;
        using (var fileStream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
        {
            string credentialPath = "token.json";
            userCredential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.FromStream(fileStream).Secrets,
                new[] { SheetsService.Scope.Spreadsheets },
                "user",
                CancellationToken.None,
                new FileDataStore(credentialPath, true));
        }

        sheetsService = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = userCredential,
            ApplicationName = applicationName,
        });
    }

    public async Task AddHeadingsAsync()
    {
        try
        {
            // Check if the sheet is blank by reading the first row
            var dataRange = $"{sheetName}!A1:Z1"; // First row only
            var sheetRequest = sheetsService.Spreadsheets.Values.Get(spreadsheetId, dataRange);
            var sheetDataValues = (await sheetRequest.ExecuteAsync()).Values;

            // If no data is present in the first row, the sheet is blank, so add headings
            if (sheetDataValues == null || sheetDataValues.Count == 0)
            {
                var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                var headers = properties.Select(p => p.Name).ToList();

                var range = $"{sheetName}!A1:{(char)('A' + headers.Count - 1)}1";
                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { headers.Cast<object>().ToList() }
                };

                var updateRequest = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                await updateRequest.ExecuteAsync();

                Console.WriteLine("Headers added to the Google Sheet.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while initializing the sheet: {ex.Message}");
        }
    }


    // Custom Add method
    public void Add(T item)
    {
        sheetData.Add(item);
        ItemAdded?.Invoke(item); // Trigger the ItemAdded event
    }

    // Event handler for when an item is added
    private void OnItemAdded(T item)
    {
        var primaryKeyProperty = GetPrimaryKeyProperty();
        var primaryKeyValue = primaryKeyProperty.GetValue(item)?.ToString();

        if (primaryKeyValue != null && !changeTracker.ContainsKey(primaryKeyValue))
        {
            changeTracker[primaryKeyValue] = item;
            rowIndexByPrimaryKey[primaryKeyValue] = sheetData.Count + 1; // Assign new row index
        }
    }

    public async Task LoadAsync()
    {
        await InitializeServiceAsync();

        await foreach (var item in ReadSheetToListAsync())
        {
            var primaryKeyProperty = GetPrimaryKeyProperty();
            var primaryKeyValue = primaryKeyProperty.GetValue(item)?.ToString();
            if (primaryKeyValue != null && !changeTracker.ContainsKey(primaryKeyValue))
            {
                changeTracker[primaryKeyValue] = item;
                rowIndexByPrimaryKey[primaryKeyValue] = sheetData.Count + 2; // Google Sheets row index starts at 1, with headers in row 1
            }
            sheetData.Add(item);
        }
    }

    public async Task SaveAsync()
    {
        var primaryKeyProperty = GetPrimaryKeyProperty();
        var dataToUpdate = new List<ValueRange>();

        foreach (var changedItem in changeTracker)
        {
            var rowData = changedItem.Value;
            var primaryKeyValue = primaryKeyProperty.GetValue(rowData)?.ToString();
            if (primaryKeyValue == null || !rowIndexByPrimaryKey.ContainsKey(primaryKeyValue)) continue;

            var rowNumber = rowIndexByPrimaryKey[primaryKeyValue];
            var updatedProperties = rowData.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var updatedValues = new List<object>();

            foreach (var property in updatedProperties)
            {
                var propertyValue = property.GetValue(rowData)?.ToString();
                updatedValues.Add(propertyValue);
            }

            var range = $"{sheetName}!A{rowNumber}:{(char)('A' + updatedProperties.Length - 1)}{rowNumber}";
            var valueRange = new ValueRange
            {
                Range = range,
                Values = new List<IList<object>> { updatedValues }
            };
            dataToUpdate.Add(valueRange);
        }

        if (dataToUpdate.Count > 0)
        {
            var batchUpdate = new BatchUpdateValuesRequest
            {
                Data = dataToUpdate,
                ValueInputOption = "USER_ENTERED"
            };
            await sheetsService.Spreadsheets.Values.BatchUpdate(batchUpdate, spreadsheetId).ExecuteAsync();
        }
    }

    private async IAsyncEnumerable<T> ReadSheetToListAsync()
    {
        var dataRange = $"{sheetName}!A1:Z";
        var sheetRequest = sheetsService.Spreadsheets.Values.Get(spreadsheetId, dataRange);
        var sheetDataValues = (await sheetRequest.ExecuteAsync()).Values;

        if (sheetDataValues?.Count > 1)
        {
            var sheetHeaders = sheetDataValues[0];
            var classProperties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var propertyMap = sheetHeaders
                .Select((header, index) => new { header, index })
                .ToDictionary(h => h.header.ToString(), h => h.index);

            int rowNumber = 2; // Starting after the header row

            foreach (var rowValues in sheetDataValues.Skip(1))
            {
                var classInstance = new T();
                string primaryKeyValue = null;
                foreach (var property in classProperties)
                {
                    if (propertyMap.TryGetValue(property.Name, out int columnIndex) && columnIndex < rowValues.Count)
                    {
                        var cellValue = rowValues[columnIndex]?.ToString();
                        if (property.IsDefined(typeof(PrimaryKeyAttribute)))
                        {
                            primaryKeyValue = cellValue; // Capture the primary key value
                        }
                        if (cellValue != null)
                        {
                            // Pass the row number and primary key value to the conversion method
                            property.SetValue(classInstance, ConvertToType(cellValue, property.PropertyType, rowNumber, primaryKeyValue));
                        }
                    }
                }
                rowNumber++; // Increment row number
                yield return classInstance;
            }
        }
    }


    private PropertyInfo GetPrimaryKeyProperty()
    {
        return typeof(T).GetProperties().FirstOrDefault(prop => Attribute.IsDefined(prop, typeof(PrimaryKeyAttribute)));
    }

    private static object ConvertToType(string value, Type targetType, int rowNumber, string primaryKeyValue)
    {
        try
        {
            var param = Expression.Parameter(typeof(string));
            var conversion = Expression.Lambda<Func<string, object>>(Expression.Convert(Expression.Call(typeof(Convert), "ChangeType", null, param, Expression.Constant(targetType)), typeof(object)), param).Compile();
            return conversion(value);
        }
        catch (Exception ex)
        {
            // Output an error message with details about the issue, including row number and primary key
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine("Error: Unable to read some data from the Google Sheet.");
            Console.WriteLine($"\nProblem: The value '{value}' in row {rowNumber} (Primary Key: {primaryKeyValue}) could not be converted to type '{targetType.Name}'. {ex.Message}");
            Console.WriteLine("\nPlease fix the incorrect data in the Google Sheet before continuing. Then stop and restart your program.");
            Console.WriteLine("\nNote: If you try to save before correcting the data, some values may be replaced with defaults.");
            Console.WriteLine("\nPress Enter to continue...");
            Console.WriteLine("------------------------------------------------");
            Console.ReadLine();
            // Return default value based on the target type
            if (targetType.IsValueType)
            {
                return Activator.CreateInstance(targetType); // Return default value for value types
            }
            return null; // Return null for reference types
        }
    }


    public IEnumerator<T> GetEnumerator()
    {
        return sheetData.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
