// ***************************************************************************************
// * BrickLink/BrickeSet Excel function integration 
// * Version 3.0 11/9/2023
// * Itamar Budin brickmindz@gmail.com
// * Using code samples from multiple resources (see internal comments for reference) 
// ***************************************************************************************
// This solution is using the Excel-DNA plug-in. If you would like more details, please look at the ExcelDna.AddIn.md file
// This version of the tool which includes
//  * Introduced new integration with BrickSet 
//  * Major code optimization to reduce code duplication (I know I can do better :)
//      
// Pre-requisites: Please make sure you follow Microsoft guidelines regarding TLS 1.2: https://support.microsoft.com/en-us/topic/applications-that-rely-on-tls-1-2-strong-encryption-experience-connectivity-failures-after-a-windows-upgrade-c46780c2-f593-8173-8670-f930816f222c
// I am not a developer, but I know how to write basic code, so please excuse any lousy code writing :)

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web;
using System.Xml;

namespace BrickLinkBrickSet
{
    public static class BricklinkBrickSetExcelIntegration
    {
        // In this section, you will need to enter the various secrets and keys that are assigned to you by Bricklink 
        // For more information see:
        //  BrickLink: https://www.bricklink.com/v3/api.page
        //  BrickSet: https://brickset.com/api/v3.asmx

        // BrickLink 
        const string consumerKey = "";                                                                  // The Consumer key
        const string consumerSecret = "";                                                               // The Consumer Secret
        const string tokenValue = "";                                                                   // The Token Value
        const string tokenSecret = "";                                                                  // The Token Secret
        const string brickLinkSetURL = "https://api.bricklink.com/api/store/v3/items/set/";             // BrickLink API Set URL
        const string brickLinkGearURL = "https://api.bricklink.com/api/store/v3/items/gear/";           // BrickLink API Gear URL
        const string brickLinkMiniFigURL = "https://api.bricklink.com/api/store/v3/items/minifig/";     // BrickLink API Minifig URL
        const string brickLinkPartURL = "https://api.bricklink.com/api/store/v3/items/part/";           // BrickLink API Part URL
        const string brickLinkBooksURL = "https://api.bricklink.com/api/store/v3/items/book/";          // BrickLink API Book URL
        const string brickLinkCategoryURL = "https://api.bricklink.com/api/store/v3/categories/";       // BrickLink API Book URL        

        // BrickSet
        const string brickSetApiKey = "";                                       // Brickset API Key
        const string bricksHash = "";                                           // Brickset Hash        
        const string brickSetSOAPUrl = "https://brickset.com/api/v3.asmx";      // Brickset URL
        const string brickSetPartNumberAttribute = "pieces";                    // Bricket set part number attribute
        const string brickSetNameAttribute = "name";                            // Bricket set name attribute
        const string brickSetYearAttribute = "year";                            // Bricket set release year attribute
        const string brickSetThemeAttribute = "theme";                          // Bricket set theme attribute
        const string brickSetImageURLAttribute = "imageURL";                    // Bricket set image attribute
        const string brickSetThumbnailURLAttribute = "thumbnailURL";            // Bricket set thumbnail attribute
        const string brickSetOriginalSellPriceAttribute = "retailPrice";        // Bricket set original sell price attribute
        const string brickSetUPCAttribute = "UPC";                              // Bricket set UPC attribute 
        const string brickSetDescriptionAttribute = "description";              // Bricket set description attribute 

        public static string tokenx = "";

        // DB Detailes (for cache)
        const string DataSource = "";                           // The DB server name
        const string InitialCatalog = "";                       // The Database NAME
        const string DBUser = "";                               // The DB username
        const string DBPassword = "";                           // The DB password        
        const string dbIDAttribute = "ID";                      // The DB column that holds the set number        
        const string dbNameAttribute = "name";                  // The DB column that holds the set name        
        const string dbTypeAttribute = "type";                  // The DB column that holds the set name        
        const string dbCategoryIDAttribute = "categoryID";      // The DB column that holds the set category  
        const string dbImageURLAttribute = "imageURL";          // The DB column that holds the set image URL 
        const string dbThumbnailURLAttribute = "thumbnail_url"; // The DB column that holds the set thumbnail URL 
        const string dbYearAttribute = "year_released";         // The DB column that holds the set release year  
        const string dbAvgPriceAttribute = "avg_price";         // The DB column that holds the set average price year (BrickLink only)
        const string dbPartNumberAttribute = "partnum";         // The DB column that holds the set part number     
        const string dbNumOfMinifigsAttribute = "minifignum";   // The DB column that holds the set number of minifig (Bricklink only)    
        const string dbUPCAttribute = "UPC";                    // The DB column that holds the set UPC    
        const string dbDescriptoinAttribute = "description";    // The DB column that holds the set description    
        const string dbOrgPriceAttribute = "original_price";    // The DB column that holds the set original price   
        const string dbSetMinifiguresAttribute = "minifigset";  // The DB columne that holds the set original price   

        // Read set infromation from DB
        private static string ReadSetInformationFromDB(String setID, String columnInput)
        {
            try
            {
                string dbResults = "N/A";
                string connectionString = @"Data Source=" + DataSource + ";Initial Catalog=" + InitialCatalog + ";User ID=" + DBUser + ";Password=" + DBPassword + ";MultipleActiveResultSets = true";
                SqlConnection dbSetConnection = new(connectionString);
                dbSetConnection.Open();
                String sql = "SELECT " + columnInput + " FROM [dbo].Sets where ID='" + setID + "'";
                SqlCommand command = new(sql, dbSetConnection);
                SqlDataReader dbSetReader = command.ExecuteReader();

                if (dbSetReader.HasRows)
                {
                    while (dbSetReader.Read())
                    {
                        dbResults = Convert.ToString(dbSetReader[columnInput]);
                    }
                    dbSetConnection.Close();
                    dbSetConnection.Close();
                }
                return dbResults;
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //**** BrickLink ********
        private static string Escape(string s)
        {
            string[] charsToEscape = new[] { "!", "*", "'", "(", ")" };
            StringBuilder escaped = new StringBuilder(Uri.EscapeDataString(s));
            foreach (var t in charsToEscape)
            {
                escaped.Replace(t, Uri.HexEscape(t[0]));
            }
            return escaped.ToString();
        }

        private static readonly string[] UriRfc3986CharsToEscape = new[] { "!", "*", "'", "(", ")" };

        // The following section was build using the example shown here: https://stackoverflow.com/questions/47378232/rest-api-authentication-oauth-1-0-using-c-sharp
        // Original Code written by https://stackoverflow.com/users/3854205/halvorsen THANK YOU!

        private static string EscapeUriDataStringRfc3986(string value)
        {
            StringBuilder escaped = new StringBuilder(Uri.EscapeDataString(value));
            for (int i = 0; i < UriRfc3986CharsToEscape.Length; i++)
            {
                escaped.Replace(UriRfc3986CharsToEscape[i], Uri.HexEscape(UriRfc3986CharsToEscape[i][0]));
            }
            return escaped.ToString();
        }

        // 11-3-2022 Added this section to try to better deal with the TIMESTAMP errors
        private static string GenerateNonce(string extra = "")
        {
            string result = "";
            SHA1 sha1 = SHA1.Create();

            Random rand = new Random();
            StringBuilder sb = new StringBuilder(1024);
            while (result.Length < 32)
            {
                sb.Length = 0;
                string[] generatedRandoms = new string[4];

                for (int i = 0; i < 4; i++)
                {
                    sb.Append(rand.Next());
                }

                sb.Append("|")
                    .Append(extra);

                result += Convert.ToBase64String(
                    sha1.ComputeHash(Encoding.ASCII.GetBytes(sb.ToString()))
                    ).TrimEnd('=')
                     .Replace("/", "")
                     .Replace("+", "");
            }

            return result.Substring(0, 32);
        }


        // This method will connect to the BrickLink API and pull the associated record including the entire payload which will help us grab individual data point like the set name, release date and average price
        // Example: 
        /*
         * https://api.bricklink.com/api/store/v3/items/set/10030-1
         * 
         *     {
                "meta": {
                    "description": "OK",
                    "message": "OK",
                    "code": 200
                },
                "data": {
                    "no": "10030-1",
                    "name": "Imperial Star Destroyer - UCS",
                    "type": "SET",
                    "category_id": 65,
                    "image_url": "//img.bricklink.com/SL/10030-1.jpg",
                    "thumbnail_url": "//img.bricklink.com/S/10030-1.gif",
                    "weight": "9093.00",
                    "dim_x": "58.80",
                    "dim_y": "50.50",
                    "dim_z": "21.00",
                    "year_released": 2002,
                    "description": "<p />Instructions for this set came in two forms.  A 'classic bound' or glued spine version, and a 'spiral bound' version.\r\n<p />The early production runs of this set featured Light Gray elements but subsequent production runs have contained varying mixtures of Light Gray and Light Bluish Gray elements.",
                    "is_obsolete": false
                }
        */
        private static string GetSetInformation(string url, string requestType)
        {
            try
            {
                // This section will determine if the original request was made using the GetSetPrice function and will change the base URL to retrieve the pricing information 
                if (requestType != "info")
                {
                    url += "/price";
                }

                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.Method = "GET";
                // Add TLS 1.2 support            
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                ServicePointManager.DefaultConnectionLimit = 9999;


                string timeStamp = ((int)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalSeconds).ToString();
                string nonce = GenerateNonce("");


                string signatureBaseString = Escape(httpWebRequest.Method.ToUpper()) + "&";
                signatureBaseString += EscapeUriDataStringRfc3986(url) + "&";
                signatureBaseString += EscapeUriDataStringRfc3986(
                    "oauth_consumer_key=" + EscapeUriDataStringRfc3986(consumerKey) + "&" +
                    "oauth_nonce=" + EscapeUriDataStringRfc3986(nonce) + "&" +
                    "oauth_signature_method=" + EscapeUriDataStringRfc3986("HMAC-SHA1") + "&" +
                    "oauth_timestamp=" + EscapeUriDataStringRfc3986(timeStamp) + "&" +
                    "oauth_token=" + EscapeUriDataStringRfc3986(tokenValue) + "&" +
                    "oauth_version=" + EscapeUriDataStringRfc3986("1.0"));

                string key = EscapeUriDataStringRfc3986(consumerSecret) + "&" + EscapeUriDataStringRfc3986(tokenSecret);
                var signatureEncoding = new ASCIIEncoding();
                byte[] keyBytes = signatureEncoding.GetBytes(key);
                byte[] signatureBaseBytes = signatureEncoding.GetBytes(signatureBaseString);
                string signatureString;
                using (var hmacsha1 = new HMACSHA1(keyBytes))
                {
                    byte[] hashBytes = hmacsha1.ComputeHash(signatureBaseBytes);
                    signatureString = Convert.ToBase64String(hashBytes);
                }
                signatureString = EscapeUriDataStringRfc3986(signatureString);

                string SimpleQuote(string s) => '"' + s + '"';

                string header = "OAuth oauth_consumer_key=" + SimpleQuote(consumerKey) + ",oauth_token=" + SimpleQuote(tokenValue) + ",oauth_signature_method=" + SimpleQuote("HMAC-SHA1") + ",oauth_timestamp=" + SimpleQuote(timeStamp) + ",oauth_nonce=" + SimpleQuote(nonce) + ",oauth_version=" + SimpleQuote("1.0") + ",oauth_signature=" + SimpleQuote(signatureString);

                httpWebRequest.Headers.Add(HttpRequestHeader.Authorization, header);
                var response = (HttpWebResponse)httpWebRequest.GetResponse();
                string characterSet = ((HttpWebResponse)response).CharacterSet;
                var responseEncoding = characterSet == ""
                    ? Encoding.UTF8
                    : Encoding.GetEncoding(characterSet ?? "utf-8");
                var responsestream = new StreamReader(response.GetResponseStream()).ReadToEnd();
                if (responsestream == null)
                {
                    throw new ArgumentNullException(nameof(characterSet));
                }
                return responsestream;
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        private static string GetSetInformationFromBrickLink(string setID, string attribute)
        {
            int setMiniFigNum = 0;
            string setData = "";
            string setInformation;
            string setMinifigureCollection = "";
            int minifigCountValue = 0;
            string typeOfRequest = "info";
            string URL = "";

            switch (attribute)
            {
                case dbNumOfMinifigsAttribute:
                    URL = "/subsets";
                    break;

                case dbSetMinifiguresAttribute:
                    URL = "/subsets";
                    break;

                case dbAvgPriceAttribute:
                    typeOfRequest = "price";
                    break;
            }

            setInformation = GetSetInformation(brickLinkSetURL + setID + "-1" + URL, typeOfRequest);

            if (setInformation.Contains("(404) Not Found") || setInformation.Contains("(400) Bad Request") || setInformation == null)
            {
                setInformation = GetSetInformation(brickLinkSetURL + setID + URL, typeOfRequest);
            }
            if (setInformation.Contains("(404) Not Found") || setInformation.Contains("(400) Bad Request") || setInformation == null)
            {
                setInformation = GetSetInformation(brickLinkGearURL + setID + URL, typeOfRequest);
            }
            // 11-3-2022 Added this section to deal with minifigure and sets who catalog ID is not                            
            if (setInformation.Contains("(404) Not Found") || setInformation.Contains("(400) Bad Request") || setInformation == null)
            {
                setInformation = GetSetInformation(brickLinkMiniFigURL + setID + URL, typeOfRequest);
            }
            if (setInformation.Contains("(404) Not Found") || setInformation.Contains("(400) Bad Request") || setInformation == null)
            {
                setInformation = GetSetInformation(brickLinkPartURL + setID + URL, typeOfRequest);
            }
            // 5-21-2023 Added this section to deal with old booklets
            if (setInformation.Contains("(404) Not Found") || setInformation.Contains("(400) Bad Request") || setInformation == null)
            {
                setInformation = GetSetInformation(brickLinkBooksURL + setID + "/subsets", typeOfRequest);
            }

            JObject setObj = JObject.Parse(setInformation);
            if (!setObj.ToString().Contains("TIMESTAMP"))
            {
                if (setObj.ContainsKey("data"))
                {
                    if (attribute == dbNumOfMinifigsAttribute)
                    {
                        IEnumerable<JToken> partsonly = setObj.SelectTokens("$..entries[?(@..type == 'MINIFIG')].quantity");

                        foreach (JToken part in partsonly)
                        {
                            setMiniFigNum += (int)part;
                        }
                        setData = setMiniFigNum.ToString();
                    }
                    else if (attribute == dbSetMinifiguresAttribute)
                    {
                        {
                            IEnumerable<JToken> minifiguresSet = setObj.SelectTokens("$..entries[?(@..type == 'MINIFIG')].item.no");
                            foreach (JToken minifigure in minifiguresSet)
                            {
                                IEnumerable<JToken> minifigureData = setObj.SelectTokens("$..entries[?(@..no == '" + minifigure + "')].quantity");
                                foreach (JToken minifigureCount in minifigureData)
                                {
                                    minifigCountValue = (int)minifigureCount;
                                }
                                setMinifigureCollection += minifigure + " (" + minifigCountValue.ToString() + "), ";
                            }

                            if (setMinifigureCollection.Length > 0)   
                                setData = setMinifigureCollection.Substring(0, setMinifigureCollection.Length - 2);
                        }
                    }
                    else if (attribute == dbCategoryIDAttribute)
                    {
                        string SetCategory = (string)setObj["data"]["category_id"] ?? "N/A";
                        if (SetCategory != "N/A")
                        {
                            var catObj = JObject.Parse(GetSetInformation(brickLinkCategoryURL + SetCategory, "info"));
                            if (!catObj.ToString().Contains("TIMESTAMP"))
                            {
                                setData = (string)catObj["data"]["category_name"] ?? "N /A";
                            }
                            else
                            {
                                // Added this sleep function as BrickLink API expect a 0.5 second delay between different call 
                                Thread.Sleep(500);
                                return GetSetCategoryFromBrickLink(setID);
                            }
                        }
                    }
                    else
                    {
                        setData = (string)setObj["data"][attribute] ?? "N/A";
                    }
                }
                return setData;
            }
            else
            {
                // Added this sleep function as BrickLink API expect a 0.5 second delay between different call 
                Thread.Sleep(500);
                return GetSetInformationFromBrickLink(setID, attribute);
            }
        }

        //This function will return the item name 
        public static string GetSetNameFromBrickLink(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setNameFromDB = ReadSetInformationFromDB(setID, dbNameAttribute);
                if (setNameFromDB == "N/A" || setNameFromDB == "" || setNameFromDB == "no results")
                {                    
                    callDB = true;
                }
                else
                {
                    return setNameFromDB;
                }

                if (callDB)
                    updateSetInCache(setID, dbNameAttribute);
                return GetSetNameFromBrickLink(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the number of minifigures for the set        
        public static string GetSetMiniFigNumberFromBrickLink(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setNumberOfMinifiguresFromDB = ReadSetInformationFromDB(setID, dbNumOfMinifigsAttribute);

                if (setNumberOfMinifiguresFromDB == "N/A" || setNumberOfMinifiguresFromDB == "" || setNumberOfMinifiguresFromDB == "no results")
                {
                    callDB = true;
                }
                else
                {
                    return setNumberOfMinifiguresFromDB;
                }
                
                if (callDB)
                    updateSetInCache(setID, dbSetMinifiguresAttribute);
                return GetSetMiniFigNumberFromBrickLink(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the number of minifigures for the set        
        public static string GetSetMiniFigCollectionFromBrickLink(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setMiniFigCollectionFromDB = ReadSetInformationFromDB(setID, dbSetMinifiguresAttribute);

                if (setMiniFigCollectionFromDB == "N/A" || setMiniFigCollectionFromDB == "" || setMiniFigCollectionFromDB == "no results")
                {
                    callDB = true;
                }
                else
                {
                    return setMiniFigCollectionFromDB;
                }
                
                if (callDB)
                    updateSetInCache(setID, dbSetMinifiguresAttribute);
                return GetSetMiniFigCollectionFromBrickLink(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item Thumbnail
        public static string GetSetThumbnailFromBrickLink(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setThumbnailFromDB = ReadSetInformationFromDB(setID, dbThumbnailURLAttribute);

                if (setThumbnailFromDB == "N/A" || setThumbnailFromDB == "" || setThumbnailFromDB == "no results")
                {
                    callDB = true;
                }
                else
                {
                    return setThumbnailFromDB;
                }
                
                if (callDB)
                    updateSetInCache(setID, dbThumbnailURLAttribute);
                return GetSetThumbnailFromBrickLink(setID);

            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item image
        public static string GetSetImageFromBrickLink(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setImageFromDB = ReadSetInformationFromDB(setID, dbImageURLAttribute);

                if (setImageFromDB == "N/A" || setImageFromDB == "" || setImageFromDB == "no results")
                {
                    callDB = true;
                }
                else
                {
                    return setImageFromDB;
                }
                
                if (callDB)
                    updateSetInCache(setID, dbImageURLAttribute);
                return GetSetImageFromBrickLink(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item release year
        public static string GetSetReleaseYearFBrickLink(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setYearFromDB = ReadSetInformationFromDB(setID, dbYearAttribute);

                if (setYearFromDB == "N/A" || setYearFromDB == "" || setYearFromDB == "no results")
                {
                    callDB = true;
                }
                else
                {
                    return setYearFromDB;
                }
                
                if (callDB)
                    updateSetInCache(setID, dbYearAttribute);
                return GetSetReleaseYearFBrickLink(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item type (Set, gear, etc.)
        public static string GetSetTypeFromBrickLink(string setID)
        {
            try
            {
                Boolean callDB = false;
                string SetTypeFromDB = ReadSetInformationFromDB(setID, dbTypeAttribute);
                if (SetTypeFromDB == "N/A" || SetTypeFromDB == "" || SetTypeFromDB == "no results")
                {
                    callDB = true;
                }
                else
                {
                    return SetTypeFromDB;
                }
                
                if (callDB)
                    updateSetInCache(setID, dbTypeAttribute);
                return GetSetTypeFromBrickLink(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item average price. Note this code is designed to pull the average price for a new set. Some older sets may not have an entry in BrinkLink so you may want to change this method
        public static string GetSetAvgPriceFromBrickLink(string setID)
        {
            try
            {
                Boolean callDB = false;
                string SetAvgPriceFromDB = ReadSetInformationFromDB(setID, dbAvgPriceAttribute);

                if (SetAvgPriceFromDB == "N/A" || SetAvgPriceFromDB == "" || SetAvgPriceFromDB == "no results")
                {
                    callDB = true;
                }
                else
                {
                    return SetAvgPriceFromDB;
                }
                
                if (callDB)
                    updateSetInCache(setID, dbAvgPriceAttribute);
                return GetSetAvgPriceFromBrickLink(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item category ID. 
        //It will compare it to the /categories API call from bricklink, and return the category name
        public static string GetSetCategoryFromBrickLink(string setID)
        {
            try
            {
                Boolean callDB = false; 
                string SetCategoryFromDB = ReadSetInformationFromDB(setID, dbCategoryIDAttribute);

                if (SetCategoryFromDB == "N/A" || SetCategoryFromDB == "" || SetCategoryFromDB == "no results")
                {
                    callDB = true;
                }
                else
                {
                    return SetCategoryFromDB;
                }

                if (callDB)
                    updateSetInCache(setID, dbCategoryIDAttribute);
                return GetSetCategoryFromBrickLink(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        public static string updateSetInCache(string setID, string dbColumn)
        // The goal for this new funtion is to streamline the cache usage of the solution. 
        // Instead of making an API call for each field in the excell file,
        // this code will first check if the the set is in the cache and if not, populate the cache for the calling function
        {
            try
            {
                string SetURL = brickLinkSetURL + setID + "-1";
                string cacheReader = ReadSetInformationFromDB(setID, dbIDAttribute);
                DateTime today = DateTime.Now;
                today.ToString("yyyy-MM-dd");
                string connectionString = @"Data Source=" + DataSource + ";Initial Catalog=" + InitialCatalog + ";User ID=" + DBUser + ";Password=" + DBPassword + ";MultipleActiveResultSets = true";
                SqlConnection setCacheConnection = new SqlConnection(connectionString);
                setCacheConnection.Open();

                string setInformation = GetSetInformation(brickLinkSetURL + setID + "-1", "info");
                if (setInformation.Contains("(404) Not Found") || setInformation.Contains("(400) Bad Request") || setInformation == null)
                {
                    setInformation = GetSetInformation(brickLinkSetURL + setID, "info");
                }
                if (setInformation.Contains("(404) Not Found") || setInformation.Contains("(400) Bad Request") || setInformation == null)
                {
                    setInformation = GetSetInformation(brickLinkGearURL + setID, "info");
                }
                // 11-3-2022 Added this section to deal with minifigure and sets who catalog ID is not                            
                if (setInformation.Contains("(404) Not Found") || setInformation.Contains("(400) Bad Request") || setInformation == null)
                {
                    setInformation = GetSetInformation(brickLinkMiniFigURL + setID, "info");
                }
                if (setInformation.Contains("(404) Not Found") || setInformation.Contains("(400) Bad Request") || setInformation == null)
                {
                    setInformation = GetSetInformation(brickLinkPartURL + setID, "info");
                }
                // 5-21-2023 Added this section to deal with old booklets
                if (setInformation.Contains("(404) Not Found") || setInformation.Contains("(400) Bad Request") || setInformation == null)
                {
                    setInformation = GetSetInformation(brickLinkBooksURL + setID, "info");
                }

                var setObj = JObject.Parse(setInformation);
                if (!setObj.ToString().Contains("TIMESTAMP"))
                {
                    // get set information
                    string setName = (string)setObj["data"]["name"] ?? "N/A";
                    string setType = (string)setObj["data"]["type"] ?? "N/A";
                    string setImageURL = "https:" + (string)setObj["data"]["image_url"] ?? "N/A";
                    string setThumbnailURL = "https:" + (string)setObj["data"]["thumbnail_url"] ?? "N/A";
                    string setYear_released = (string)setObj["data"]["year_released"] ?? "N/A";

                    // Get set price
                    string setAVGPrice = GetSetInformationFromBrickLink(setID, dbAvgPriceAttribute);

                    // get set category
                    string setCategory_id = GetSetInformationFromBrickLink(setID, dbCategoryIDAttribute);

                    // get set number of parts from brickset
                    string setPartNum = GetSetAttributeFromBrickSet(setID, brickSetPartNumberAttribute);
                    if (setPartNum == "BrickSet API limit exceeded")
                        setPartNum = "0";

                    // get set number of minifigures
                    string setMinifigureNum = GetSetInformationFromBrickLink(setID, dbNumOfMinifigsAttribute);

                    // get set UPC
                    string setUPC = GetSetAttributeFromBrickSet(setID, brickSetUPCAttribute);
                    if (setUPC == "BrickSet API limit exceeded")
                        setUPC = "0";

                    // get set description
                    string setDescription = GetSetAttributeFromBrickSet(setID, brickSetDescriptionAttribute);
                    if (setDescription == "BrickSet API limit exceeded")
                        setDescription = "";

                    // get set orignal retail price
                    string setOriginalPrice = GetSetAttributeFromBrickSet(setID, brickSetOriginalSellPriceAttribute + "US");
                    if (setOriginalPrice == "BrickSet API limit exceeded")
                        setOriginalPrice = "";

                    // get set minifigure collection
                    string setMinifiguresCollection = GetSetInformationFromBrickLink(setID, dbSetMinifiguresAttribute);

                    if (cacheReader != setID)
                    {
                        string insertsql = "INSERT INTO dbo.Sets (ID,name,type,categoryID,imageURL,thumbnail_url,year_released,avg_price,date_updated,partnum,minifignum,UPC,description,original_price,minifigset) VALUES (@ID,@name,@type,@categoryID,@imageURL,@thumbnail_url,@year_released,@avg_price,@date_updated,@partnum,@minifignum,@UPC,@description,@original_price,@minifigset)";
                        SqlCommand insertCommand = new SqlCommand(insertsql, setCacheConnection);
                        insertCommand.Parameters.AddWithValue("@ID", setID);
                        insertCommand.Parameters.AddWithValue("@name", WebUtility.HtmlDecode(setName));
                        insertCommand.Parameters.AddWithValue("@type", setType);
                        insertCommand.Parameters.AddWithValue("@categoryID", WebUtility.HtmlDecode(setCategory_id));
                        insertCommand.Parameters.AddWithValue("@imageURL", setImageURL);
                        insertCommand.Parameters.AddWithValue("@thumbnail_url", setThumbnailURL);
                        insertCommand.Parameters.AddWithValue("@year_released", setYear_released);
                        insertCommand.Parameters.AddWithValue("@avg_price", setAVGPrice);
                        insertCommand.Parameters.AddWithValue("@date_updated", today);
                        insertCommand.Parameters.AddWithValue("@partnum", int.Parse(setPartNum));
                        insertCommand.Parameters.AddWithValue("@minifignum", int.Parse(setMinifigureNum));
                        insertCommand.Parameters.AddWithValue("@UPC", setUPC);
                        insertCommand.Parameters.AddWithValue("@description", setDescription);
                        insertCommand.Parameters.AddWithValue("@original_price", setOriginalPrice);
                        insertCommand.Parameters.AddWithValue("@minifigset", setMinifiguresCollection);
                        string debugSQL = insertCommand.ToString();
                        int result = insertCommand.ExecuteNonQuery();

                        // check for errors
                        if (result < 0)
                        {
                            setCacheConnection.Close();
                            return "There was an error inserting the values to the DB";
                        }
                        return "Entry was added to the DB";
                    }
                    else
                    {
                        if (setName != "N/A")
                        {
                            string updatesql = "update dbo.Sets set name=@name,type=@type,categoryID=@category_id,imageURL=@imageURL,thumbnail_url=@thumbnail_url,year_released=@year_released,avg_price=@avg_price,date_updated=@date_updated,partnum=@partnum,minifignum=@minifignum,UPC=@UPC,description=@description,original_price=@original_price,minifigset=@minifigset where ID=@ID";
                            SqlCommand updateCommand = new SqlCommand(updatesql, setCacheConnection);
                            updateCommand.Parameters.AddWithValue("@ID", setID);
                            updateCommand.Parameters.AddWithValue("@name", HttpUtility.HtmlDecode(setName));
                            updateCommand.Parameters.AddWithValue("@type", setType);
                            updateCommand.Parameters.AddWithValue("@category_id", HttpUtility.HtmlDecode(setCategory_id));
                            updateCommand.Parameters.AddWithValue("@imageURL", setImageURL);
                            updateCommand.Parameters.AddWithValue("@thumbnail_url", setThumbnailURL);
                            updateCommand.Parameters.AddWithValue("@year_released", setYear_released);
                            updateCommand.Parameters.AddWithValue("@avg_price", setAVGPrice);
                            updateCommand.Parameters.AddWithValue("@date_updated", today);
                            updateCommand.Parameters.AddWithValue("@partnum", int.Parse(setPartNum));
                            updateCommand.Parameters.AddWithValue("@minifignum", int.Parse(setMinifigureNum));
                            updateCommand.Parameters.AddWithValue("@UPC", setUPC);
                            updateCommand.Parameters.AddWithValue("@description", setDescription);
                            updateCommand.Parameters.AddWithValue("@original_price", setOriginalPrice);
                            updateCommand.Parameters.AddWithValue("@minifigset", setMinifiguresCollection);
                            int result = updateCommand.ExecuteNonQuery();

                            // check for errors
                            if (result < 0)
                            {
                                setCacheConnection.Close();
                                return "There was an error updating the values to the DB";
                            }
                        }
                        return "DB Record updated";
                    }
                }
                else
                {
                    // Added this sleep function as BrickLink API expect a 0.5 second delay between different call 
                    Thread.Sleep(500);
                    return updateSetInCache(setID, dbColumn);
                }
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //***** BrickSet ******

        // Calling the BrickSet API
        private static HttpWebRequest CreateSOAPWebRequest()
        {
            //Making Web Request    
            HttpWebRequest Req = (HttpWebRequest)WebRequest.Create(@brickSetSOAPUrl);
            //Content_type    
            Req.ContentType = "application/soap+xml;charset=utf-8";
            //HTTP method    
            Req.Method = "POST";
            //return HttpWebRequest    
            return Req;
        }

        private static string GetSetInformationFromBrickSet(string SetID)
        {
            try
            {
                //Calling CreateSOAPWebRequest method  
                var ServiceResult = "";
                var soapNs = @"http://www.w3.org/2003/05/soap-envelope";
                var brickSetNs = @"https://brickset.com/api/";


                HttpWebRequest request = CreateSOAPWebRequest();

                //SOAP Body Request    
                var doc = new XmlDocument();
                var root = doc.AppendChild(doc.CreateElement("soap12", "Envelope", soapNs));
                doc.DocumentElement.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
                doc.DocumentElement.SetAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
                XmlDeclaration xml_declaration;
                xml_declaration = doc.CreateXmlDeclaration("1.0", "utf-8", "yes");
                XmlElement document_element = doc.DocumentElement;
                doc.InsertBefore(xml_declaration, document_element);
                var body = root.AppendChild(doc.CreateElement("soap12", "Body", soapNs));
                var getSets = body.AppendChild(doc.CreateElement("api", "getSets", brickSetNs));

                getSets.AppendChild(doc.CreateElement("api", "apiKey", brickSetNs)).InnerText = brickSetApiKey;
                getSets.AppendChild(doc.CreateElement("api", "userHash", brickSetNs)).InnerText = bricksHash;
                getSets.AppendChild(doc.CreateElement("api", "params", brickSetNs)).InnerText = "{setNumber:'" + SetID + "-1'}";

                using (Stream stream = request.GetRequestStream())
                {
                    doc.Save(stream);
                }
                //Geting response from request    
                using WebResponse Serviceres = request.GetResponse();
                using (StreamReader rd = new(Serviceres.GetResponseStream()))
                {
                    //reading stream    
                    ServiceResult = rd.ReadToEnd();
                }
                return ServiceResult;
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function handles the logic to read information from BrickSet
        private static string GetSetAttributeFromBrickSet(string setID, string attribute)
        {
            try
            {
                string setInformation = "";
                var SOAPResponse = GetSetInformationFromBrickSet(setID).ToString();
                int index = SOAPResponse.IndexOf("<?xml");
                String cleanJSONPayload = SOAPResponse.Substring(0, index);
                JObject setObj = JObject.Parse(cleanJSONPayload);
                if (!setObj.ToString().Contains("API limit exceeded"))
                {
                    if (attribute == brickSetImageURLAttribute || attribute == brickSetThumbnailURLAttribute)
                    {
                        setInformation = (string)setObj["sets"][0]["image"][attribute] ?? "N/A";
                    }
                    else if (attribute == brickSetUPCAttribute)
                    {
                        setInformation = (string)setObj["sets"][0]["barcode"][attribute] ?? "N/A";
                        if (setInformation == "N/A")
                            setInformation = (string)setObj["sets"][0]["barcode"]["EAN"] ?? "N/A";
                    }
                    else if (attribute == brickSetDescriptionAttribute)
                    {
                        setInformation = HttpUtility.HtmlDecode((string)setObj["sets"][0]["extendedData"][attribute]).Replace("<p>", "").Replace("</p>", "") ?? "N/A";
                    }
                    else if (attribute.Contains(brickSetOriginalSellPriceAttribute))
                    {
                        string countryCode = attribute.Substring(11);
                        setInformation = (string)setObj["sets"][0]["LEGOCom"][countryCode][brickSetOriginalSellPriceAttribute] ?? "N/A";
                    }
                    else
                    {
                        setInformation = (string)setObj["sets"][0][attribute] ?? "N/A";
                    }
                    return setInformation;
                }
                else
                {
                    return "BrickSet API limit exceeded";
                }
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item name 
        public static string GetSetNameFromBrickSet(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setNameFromDB = ReadSetInformationFromDB(setID, dbNameAttribute);
                if (setNameFromDB == "N/A" || setNameFromDB == "" || setNameFromDB == "no results")
                {
                    if (setID == "BrickSet API limit exceeded")
                    {
                        return setID;
                    }
                    else
                    {
                        callDB = true;
                    }
                }                
                
                if (callDB)
                    updateSetInCache(setID, dbNameAttribute);
                return GetSetNameFromBrickSet(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item theme 
        public static string GetSetThemeFromBrickSet(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setPartNumFromDB = ReadSetInformationFromDB(setID, dbCategoryIDAttribute);
                {
                    if (setPartNumFromDB == "N/A" || setPartNumFromDB == "" || setPartNumFromDB == "no results")
                    {
                        if (setID == "BrickSet API limit exceeded")
                        {
                            return setID;
                        }
                        else 
                        {
                            callDB = true;
                        }                        
                    }
                    else
                    {
                        return setPartNumFromDB;
                    }
                }
                
                if (callDB)
                    updateSetInCache(setID, dbTypeAttribute);
                return GetSetThemeFromBrickSet(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item image URL 
        public static string GetSetImageURLFromBrickSet(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setImageURLFromDB = ReadSetInformationFromDB(setID, dbImageURLAttribute);
                if (setImageURLFromDB == "N/A" || setImageURLFromDB == "" || setImageURLFromDB == "no results")
                {
                    if (setID == "BrickSet API limit exceeded")
                    {
                        return setID;
                    }
                    else
                    {
                        callDB = true;
                    }
                }
                else
                {
                    return setImageURLFromDB;
                }
                
                if (callDB)
                    updateSetInCache(setID, dbImageURLAttribute);
                return GetSetThemeFromBrickSet(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item thumbnail URL 
        public static string GetSetThumbnailURLFromBrickSet(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setThumbnailURLFromDB = ReadSetInformationFromDB(setID, dbThumbnailURLAttribute);
                {
                    if (setThumbnailURLFromDB == "N/A" || setThumbnailURLFromDB == "" || setThumbnailURLFromDB == "no results")
                    {
                        if (setID == "BrickSet API limit exceeded")
                        {
                            return setID;
                        }
                        else
                        {
                            callDB = true;
                        }
                    }
                    else
                    {
                        return setThumbnailURLFromDB;
                    }
                }
                if (callDB)
                    updateSetInCache(setID, dbThumbnailURLAttribute);
                return GetSetThumbnailURLFromBrickSet(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item release year 
        public static string GetSetReleaseYearFromBrickSet(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setReleaseYearFromDB = ReadSetInformationFromDB(setID, dbYearAttribute);

                if (setReleaseYearFromDB == "N/A" || setReleaseYearFromDB == "" || setReleaseYearFromDB == "no results")
                {
                    if (setID == "BrickSet API limit exceeded")
                    {
                        return setID;
                    }
                    else
                    {
                        callDB = true;
                    }
                }
                else
                {
                    return setReleaseYearFromDB;
                }
                
                if (callDB)
                    updateSetInCache(setID, dbYearAttribute);
                return GetSetReleaseYearFromBrickSet(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item part number
        public static string GetSetPartsNumberFromBrickSet(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setPartNumFromDB = ReadSetInformationFromDB(setID, dbPartNumberAttribute);
                if (setPartNumFromDB == "N/A" || setPartNumFromDB == "" || setPartNumFromDB == "no results")
                {
                    if (setID == "BrickSet API limit exceeded")
                    {
                        return setID;
                    }
                    else
                    {
                        callDB = true;
                    }

                }
                else
                {
                    return setPartNumFromDB;
                }
                
                if (callDB)
                    updateSetInCache(setID, dbPartNumberAttribute);
                return GetSetPartsNumberFromBrickSet(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item UPC 
        public static string GetSetUPCFromBrickSet(string setID)
        {
            try
            {
                Boolean callDB = false;
                string setUPCFromDB = ReadSetInformationFromDB(setID, dbUPCAttribute);

                if (setUPCFromDB == "N/A" || setUPCFromDB == "" || setUPCFromDB == "no results")
                {
                    if (setID == "BrickSet API limit exceeded")
                    {
                        return setID;
                    }
                    else
                    {
                        callDB = true;
                    }

                }
                else
                {
                    return setUPCFromDB;

                }
                
                if (callDB)
                    updateSetInCache(setID, dbUPCAttribute);
                return GetSetUPCFromBrickSet(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item description 
        public static string GetSetDescriptionFromBrickSet(string setID)
        {
            try
            {
                Boolean callDB = false;
                string SetDescriptionFromDB = ReadSetInformationFromDB(setID, dbDescriptoinAttribute);

                if (SetDescriptionFromDB == "N/A" || SetDescriptionFromDB == "" || SetDescriptionFromDB == "no results")
                {
                    if (setID == "BrickSet API limit exceeded")
                    {
                        return setID;
                    }
                    else
                        {
                        callDB = true;
                    }

                }
                else
                {
                    return SetDescriptionFromDB;
                }
                
                if (callDB)
                    updateSetInCache(setID, dbDescriptoinAttribute);
                return GetSetDescriptionFromBrickSet(setID);
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item original retail price
        public static string GetSetOriginalPriceFromBrickSet(string setID, string country)
        {
            try
            {
                if (country == "US" || country == "UK" || country == "DE" || country == "CA")
                {
                    Boolean callDB = false;
                    string setOrgPriceFromDB = ReadSetInformationFromDB(setID, dbOrgPriceAttribute);

                    if (setOrgPriceFromDB == "N/A" || setOrgPriceFromDB == "" || setOrgPriceFromDB == "no results")
                    {
                        if (setID == "BrickSet API limit exceeded")
                        {
                            return setID;
                        }
                        else
                        {
                            callDB = true;
                        }

                    }
                    else
                    {
                        return setOrgPriceFromDB;
                    }
                    
                    if (callDB)
                        updateSetInCache(setID, dbOrgPriceAttribute);
                    return GetSetOriginalPriceFromBrickSet(setID, country);
                }
                else
                    return "Wrong country code entered, accetable values are: US, UK, CA and DE";
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        public static void Main()
        {
        }
    }
}
