# Google Sheets API Integration Example

This program demonstrates how to use the Google Sheets API to interact with a Google Sheets spreadsheet that stores pet data. It includes functionality for adding, updating, and retrieving pet records, using Google Sheets as the data store.

## Features

- Load pet data from a Google Sheet into the application.
- Add new pet records (e.g., name, species, age, price, color).
- Update existing pet records.
- Save changes back to the Google Sheet.
- Display all pet records in the console.

## Prerequisites

Before running the program, make sure you have completed the following steps:

1. **Google Cloud Project:**
   - Create a project in the [Google Cloud Console](https://console.cloud.google.com/).
   - Enable the **Google Sheets API** for the project.
   - Set up **OAuth 2.0** credentials for a desktop application.
   - Download the `credentials.json` file and place it in your project directory.

2. **Google Sheet:**
   - Create a Google Sheet and rename the first sheet (e.g., `Pets`).
   - Add the following columns as headers: `Name`, `Species`, `Age`, `Price`, `Color`.
   - Note down the **Sheet ID** from the URL.

3. **NuGet Package:**
   - Install the **Google Sheets API Client Library** by running the following command in the **Package Manager Console**:
     ```
     Install-Package Google.Apis.Sheets.v4
     ```

## Setup Instructions

1. **Google Sheets API Configuration:**
   - Follow the instructions at the top of the `Program.cs` file to set up the Google Sheets API and OAuth 2.0.
   - Ensure your `credentials.json` file is placed in the project directory.
   - **Important**: Add `credentials.json` to your `.gitignore` file to prevent it from being committed to source control.

2. **Project Directory Structure:**
   - Place the `credentials.json` file in the project directory.
   - Ensure the `Program.cs` file is located in the `src` directory of your solution.

3. **Running the Program:**
   - The program loads pet data from the Google Sheet, checks if certain pets (e.g., "Fluffy" and "Buddy") exist, and adds them if they don't.
   - It also increments the age of "Buddy" and saves the changes back to the Google Sheet.
   - The current list of pets is printed in the console, displaying their names, species, ages, prices, and colors.

## Usage

### Initialize and Load Data:
```csharp
var sheet = new GoogleSheet<Pet>("Pet Shop", "google_sheet_id_here", "Pets");
await sheet.LoadAsync();
```

### Add a New Pet:
```csharp
var newPet = new Pet
{
    Name = "Fluffy",
    Species = "Cat",
    Age = 3,
    Price = 100,
    Color = "White"
};
sheet.Add(newPet);
await sheet.SaveAsync();
```

### Update an Existing Pet:
```csharp
var buddy = sheet.FirstOrDefault(pet => pet.Name == "Buddy");
if (buddy != null)
{
    buddy.Age += 1;
}
await sheet.SaveAsync();
```

### List All Pets:
```csharp
foreach (var pet in sheet)
{
    Console.WriteLine($"{pet.Name}\t{pet.Species}\t{pet.Age}\t{pet.Price}\t{pet.Color}");
}
```
