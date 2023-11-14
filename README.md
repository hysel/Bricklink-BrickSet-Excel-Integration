# Bricklink/BrickSet Excel Integration
This project aims to assist you in accessing information from LEGO sets, including Minifigures and other items, using Excel functions and data storage in an Excel spreadsheet. The code can retrieve the name, release year, and average price for new items directly from the BrickLink or BricketSet APIs. I have been working on this project for more than two years, and since I am not a developer, I would like to invite you to improve this solution by introducing additional features beyond the current scope.

**Prerequisites:**


1) Install Visual Studio 2022 https://visualstudio.microsoft.com/downloads/
2) Install SQL Server (needed for reporting) https://www.microsoft.com/en-us/sql-server/sql-server-downloads
3) Install SQL Server Management Studio (SSMS) https://learn.microsoft.com/en-us/sql/ssms/download-sql-server-management-studio-ssms?view=sql-server-ver16#download-ssms
4) Follow this link to ensure TLS 1.2 is enabled on your machine.

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

**Preparing the code:**

1) Obtain the four secret keys to connect to BrickLink API. (For more details, see: https://www.bricklink.com/v2/api/welcome.page)
2) Obtain your BrickSet API key (For more details, see: https://brickset.com/article/52664/api-version-3-documentation
3) Once you have your API Key, get your unique Hash by calling the Login method (https://brickset.com/api/v3.asmx?op=login)
4) If you are running Windows 10 and above, you will need to add a special registry key to your OS to allow support TLS 1.2 (https://support.microsoft.com/en-us/topic/applications-that-rely-on-tls-1-2-strong-encryption-experience-connectivity-failures-after-a-windows-upgrade-c46780c2-f593-8173-8670-f930816f222c)
5) Open the solution in Visual Studio and update the following attributes:
        
        const string brickSetApiKey = "";    // Brickset API Key

        const string bricksHash = "";        // Brickset Hash
   
        const string consumerKey = "";        // The Consumer key
        
        const string consumerSecret = "";     // The Consumer Secret
        
        const string tokenValue = "";         // The Token Value
        
        const string tokenSecret = "";        // The Token Secret               
        
        const string DataSource = "";         // The DB server name
        
        const string InitialCatalog = "";     // The Database NAME
        
        const string DBUser = "";             // The DB username
        
        const string DBPassword = "";         // The DB password
        

7) Save the File
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
- GetSetOriginalPriceFromBrickSet - This function will return the item's original retail price
- 
**Miscellaneous:**
updateSetInCache - This function will update the item record in the DB

**Known Limitations:**
- BrickSet limits the number of API Calls, so please ensure you have the proper permission before using it.
- To avoid a large number of API loads, copy and paste the values of the queries back to the Excel spreadsheet.


