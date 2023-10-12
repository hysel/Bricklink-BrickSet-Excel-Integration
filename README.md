# Bricklink-Excel-Integration
This project aims to assist you in accessing information from LEGO sets, including Minifigures and other items, with the aid of Excel functions and data storage in an Excel spreadsheet. The code has the ability to retrieve the name, release year, and average price for new items directly from the BrickLink API. Our team has been working on this project for two years, and since I am not a developer, I invite you to improve this solution by introducing additional features beyond the current scope.

Usage:

1) Obtain the four secret keys to connect to BrickLink API. (For more details, see: https://www.bricklink.com/v2/api/welcome.page)

2) If you are running Windows 10 and above, you will need to add a special registry key to your OS to allow support TLS 1.2 (https://support.microsoft.com/en-us/topic/applications-that-rely-on-tls-1-2-strong-encryption-experience-connectivity-failures-after-a-windows-upgrade-c46780c2-f593-8173-8670-f930816f222c) 

3) Open the solution in Visual Studio and update the following attributes:
        
        const string consumerKey = "";        // The Consumer key
        
        const string consumerSecret = "";     // The Consumer Secret
        
        const string tokenValue = "";         // The Token Value
        
        const string tokenSecret = "";        // The Token Secret               
        
        const string DataSource = "";         // The DB server name
        
        const string InitialCatalog = "";     // The Database NAME
        
        const string DBUser = "";             // The DB username
        
        const string DBPassword = "";         // The DB password
        

4) Compile the code.

5) Go to the solution folder (for example C:\Users\<user name>\source\repos\Bricklink-Excel-Integration\bin\Debug) and run the BrickLink-AddIn64.xll file.

6) Enable all Macros support in MS Excel.

7) Create a new sheet and enter the set number in any field.

8) Go to another field, and you will be able to call the following functions:

    - GetSetName(<set Number>) - Get the LEGO set Name
    - GetSetThumbnail(<set Number>)- Get the LEGO set Thumbnail URL
    - GetSetImage(<set Number>)- Get the LEGO set Image URL
    - GetSetYear(<set Number>)- Get the LEGO set release date
    - GetSetPrice(<set Number>)- Get the LEGO set average price. (Yes, I know this is not a good idea, as the information is skewed due to someone putting ridiculous prices on a set. I will fix it in a later date)
    - GetSetCategory(<set Number>) - Get the LEGO set category. This is based on an XML file that is part of the solution that translates BrickLink category ID (found on the set JSON payload) to the human-friendly Name.
    - GetSetPartsNumber(<set Number>) - This method will return the total number of parts, including extras for the set
    - GetSetMinifigNumber(<set Number>) - This method will return the number of minifigures for the set  


