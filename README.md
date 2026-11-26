# Bricklink/BrickSet Excel Integration
This project aims to assist you in accessing information from LEGO sets, including Minifigures and other items, using Excel functions and data storage in an Excel spreadsheet. The code can retrieve the name, release year, and average price for new items directly from the BrickLink or BricketSet APIs. I have been working on this project for more than two years, and since I am not a developer, I invite you to improve this solution by introducing additional features beyond the current scope.

**Prerequisites:**


1) Install Visual Studio 2022 https://visualstudio.microsoft.com/downloads/
2) Install SQL Server (needed for reporting) https://www.microsoft.com/en-us/sql-server/sql-server-downloads
3) Install SQL Server Management Studio (SSMS) https://learn.microsoft.com/en-us/sql/ssms/download-sql-server-management-studio-ssms?view=sql-server-ver16#download-ssms
4) Follow this https://support.microsoft.com/en-us/topic/applications-that-rely-on-tls-1-2-strong-encryption-experience-connectivity-failures-after-a-windows-upgrade-c46780c2-f593-8173-8670-f930816f222c to ensure TLS 1.2 is enabled on your machine.

**Loading the Code:**
1) Launch Visual Studio
2) Click on “Clone a repository.”
3) In the Repository Location, type: https://github.com/hysel/Bricklink-Excel-Integration
4) Click on the Clone button.

**Setting up the DB:**

1) Launch SQL Server Management Studio (SSMS)
2) Connect to your local DB server
3) Create a new DB name, BrickLinkCache
4) Create a local SQL server user and assign owner permissions on the created DB.
5) Go back to the code
6) Locate the “Create_DB_Table.sql” file
7) Copy the content of the file.
8) Back in SSMS, open a new query window and paste the value of the SQL file
9) Run the file
10) Make sure that the Sets table is created

**Existing legacy cache:** Duplicate set IDs are supported. The add-in reads the newest matching row by `date_updated` and refreshes one matching row in place. It keeps the older rows in the cache. The existing schema can be used as-is.

**Preparing the code:**

1) Obtain the four secret keys to connect to BrickLink API. (For more details, see: https://www.bricklink.com/v2/api/welcome.page)
2) Obtain your BrickSet API key (For more details, see: https://brickset.com/article/52664/api-version-3-documentation
3) Once you have your API Key, get your unique Hash by calling the Login method (https://brickset.com/api/v3.asmx?op=login)
4) The legacy add-in uses TLS 1.2 for BrickLink and BrickSet HTTPS connections. It configures its own HTTP client and uses Windows certificate validation.
5) Run [Set-LegacyCredentials.ps1](Set-LegacyCredentials.ps1) from the repository root under the Windows account that runs Excel.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Set-LegacyCredentials.ps1
```

The menu shows whether each entry is set or missing without displaying its value. Choose one or more numbers, separated by commas, or press Enter to select all entries. The script masks the credential prompts. Press Enter on an empty credential prompt to keep the existing entry. After each update, the menu reopens and refreshes the status. Choose **0. Exit** when finished. The execution-policy override applies only to the PowerShell process started by this command.

Choose **11. Test Connections** to run a read-only BrickLink catalog request, validate the BrickSet API key and user hash, and check SQL access to the cache table without retrieving rows. The tests run in PowerShell. The menu displays only each service's status. You can run the tests without building the add-in.

For manual setup, open **Credential Manager → Windows Credentials → Add a generic credential**. Create one entry per row. Enter the target under **Internet or network address**, enter `BrickLinkBrickSet` under **User name**, and put the actual value in **Password** for every entry.

| Target | Value for the Password field |
| --- | --- |
| `BrickLinkBrickSet/Legacy/BrickLinkConsumerKey` | BrickLink consumer key |
| `BrickLinkBrickSet/Legacy/BrickLinkConsumerSecret` | BrickLink consumer secret |
| `BrickLinkBrickSet/Legacy/BrickLinkToken` | BrickLink token |
| `BrickLinkBrickSet/Legacy/BrickLinkTokenSecret` | BrickLink token secret |
| `BrickLinkBrickSet/Legacy/BrickSetApiKey` | BrickSet API key |
| `BrickLinkBrickSet/Legacy/BrickSetUserHash` | BrickSet user hash |
| `BrickLinkBrickSet/Legacy/SqlUser` | SQL cache username |
| `BrickLinkBrickSet/Legacy/SqlPassword` | SQL cache password |
| `BrickLinkBrickSet/Legacy/SqlServer` | SQL Server name or instance |
| `BrickLinkBrickSet/Legacy/SqlDatabase` | SQL cache database name |

6) Load the add-in from the same Windows account. To change credentials, the server, or the database, edit the relevant entries and recalculate the formulas. The add-in reads the current values when a connection is needed. If you previously configured environment variables for this add-in, remove them from your Windows user settings.

7) To build and debug the legacy add-in, open the solution in Visual Studio.
8) Press ALT+F7 to open the project properties
9) Under Debug, select the "Start external program" option
10) enter the path to your excel.exe file (C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE)
11) Under "Start Options," make sure the "Command line arguments" has the value of '/x "BrickLinkBrickSet-AddIn64.xll'
12) Compile the code.
13) Go to the solution folder (for example C:\Users\<user name>\source\repos\Bricklink-BrickSet-Excel-Integration\bin\Debug) and run the either BrickLinkBrickSet-AddIn64.xll for x64 office installtions or the BrickLinkBrickSet-AddIn.xll for the 32Bit version.
14) Enable all Macros support in MS Excel.
15) Create a new sheet and enter the set number in any field (Note that the code expects the set/item number or designation as they are shown in BrickLink or BrickSet)
16) Go to another field, and you will be able to call the following functions while referencing the original field)

**BrickLink:**
- GetSetNameFromBrickSet - This function will return the item name
- GetSetMiniFigNumberFromBrickLink - This function will return the number of minifigures for the set
- GetSetMiniFigCollectionFromBrickLink - This function will return the number of minifigures for the set
- GetSetThumbnailFromBrickLink - This function will return the item Thumbnail
- GetSetImageFromBrickLink - This function will return the item image
- GetSetReleaseYearFBrickLink - This function will return the item release year
- GetSetTypeFromBrickLink - This function will return the item type (Set, gear, etc.)
- GetSetAvgPriceFromBrickLink - This function will return the item's average price.
- GetSetCategoryFromBrickLink - This function will return the item category ID.

**BrickSet:**
- GetSetNameFromBrickSet - This function will return the item name
- GetSetThemeFromBrickSet - This function will return the item theme
- GetSetImageURLFromBrickSet - This function will return the item image URL
- GetSetThumbnailURLFromBrickSet - This function will return the item thumbnail URL
- GetSetReleaseYearFromBrickSet - This function will return the item release year
- GetSetPartsNumberFromBrickSet - This function will return the item part number
- GetSetUPCFromBrickSet - This function will return the item UPC
- GetSetDescriptionFromBrickSet - This function will return the item description
- GetSetOriginalPriceFromBrickSet - This function will return the item's original retail price. Acceptable input: US, UK, CA, DE
  
**Miscellaneous:**
**Legacy cache and parallel calculation**

- The worksheet lookups can participate in Excel's multithreaded calculation. Simultaneous misses for the same set share an in-flight refresh. Overlapping requests to BrickSet for the same set number also share one response. Up to four provider lookups run at once. SQL checks and writes each set inside a short transaction after the provider calls finish.
- The cache stores US retail prices. UK, DE, and CA requests use the corresponding BrickSet market directly. A missing regional price returns `N/A`.
- Missing image links return `N/A` if neither provider supplies a usable URL. Minifigure collections list each ID once with its total quantity.
- To replace an older cached record, explicitly refresh it with `updateSetInCache(setID, dbColumn)`, then recalculate. The second argument is retained for compatibility. A refresh writes the whole record.

**Known Limitations:**
- BrickSet limits the number of API Calls, so please ensure you have the proper permission before using it.
- To avoid a large number of API loads, copy and paste the values of the queries back to the Excel spreadsheet.


