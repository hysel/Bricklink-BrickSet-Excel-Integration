// ***************************************************************************************
// * BrickLink Excel function integration 
// * Version 2.0 10/31/2022
// * Itamar Budin brickmindz@gmail.com
// * Using code samples from multiple resource (see internal comments for reference) 
// ***************************************************************************************
// This solution is using the Excel-DNA plug-in. For more details, see the ExcelDna.AddIn.md file
// This is version of the tool which includes
//  * A new DB based cache solution
//  * Update to support BrickLink v2 API that was published in 10/31/2022
//  * Variables normalization
// Pre-requisits: Please make sure you follow Microsoft guidelines regarding TLS 1.2: https://support.microsoft.com/en-us/topic/applications-that-rely-on-tls-1-2-strong-encryption-experience-connectivity-failures-after-a-windows-upgrade-c46780c2-f593-8173-8670-f930816f222c

// I am not a developer but know how to write basic code so please excuse any bad code writing :)




using Newtonsoft.Json.Linq;
using System;
using System.Data;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Xml;
using System.Data.SqlClient;
using System.Web;

namespace BrickLink
{
    public static class BricklinkExcelIntegration
    {
        // TODO: In this section, you will need to enter the various secrets and keys that are assigned to you by Bricklink
        // For more information see: https://www.bricklink.com/v3/api.page

        const string consumerKey = "";        // The Consumer key
        const string consumerSecret = "";     // The Consumer Secret
        const string tokenValue = "";         // The Token Value
        const string tokenSecret = "";        // The Token Secret
        const string brickLinkURL = "https://api.bricklink.com/api/store/v2/items/set/"; // BrickLink API URL
        public static string tokenx = "";

        // Added a new section to implmenet DB based cache
        const string DataSource = "";       // The DB server name
        const string InitialCatalog = "";   // The Database NAME
        const string DBUser = "";           // The DB username
        const string DBPassword = "";       // The DB password
   

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

        // This method will connect to the BrickLink API and pull the associated record including the entire payload which will help us grab individual data point like the set name, release date and average price
        // Example: 
        /*
         * https://api.bricklink.com/api/store/v2/items/set/10030-1
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
        public static string GetSetInformation(string url, string requestType)
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
            string nonce = Convert.ToBase64String(Encoding.UTF8.GetBytes(timeStamp));


            string signatureBaseString = Escape(httpWebRequest.Method.ToUpper()) + "&";
            signatureBaseString += EscapeUriDataStringRfc3986(url) + "&";
            signatureBaseString += EscapeUriDataStringRfc3986(
                "oauth_consumer_key=" + EscapeUriDataStringRfc3986(consumerKey) + "&" +
                "oauth_nonce=" + EscapeUriDataStringRfc3986(nonce) + "&" +
                "oauth_signature_method=" + EscapeUriDataStringRfc3986("HMAC-SHA1") + "&" +
                "oauth_timestamp=" + EscapeUriDataStringRfc3986(timeStamp) + "&" +
                "oauth_token=" + EscapeUriDataStringRfc3986(tokenValue) + "&" +
                "oauth_version=" + EscapeUriDataStringRfc3986("1.0"));
            Console.WriteLine(@"signatureBaseString: " + signatureBaseString);

            string key = EscapeUriDataStringRfc3986(consumerSecret) + "&" + EscapeUriDataStringRfc3986(tokenSecret);
            Console.WriteLine(@"key: " + key);
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
            Console.WriteLine(@"signatureString: " + signatureString);

            string SimpleQuote(string s) => '"' + s + '"';

            string header = "OAuth oauth_consumer_key=" + SimpleQuote(consumerKey) + ",oauth_token=" + SimpleQuote(tokenValue) + ",oauth_signature_method=" + SimpleQuote("HMAC-SHA1") + ",oauth_timestamp=" + SimpleQuote(timeStamp) + ",oauth_nonce=" + SimpleQuote(nonce) + ",oauth_version=" + SimpleQuote("1.0") + ",oauth_signature=" + SimpleQuote(signatureString);

            Console.WriteLine(@"header: " + header);
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

        //This function will return the item name
        public static string GetSetName(string setID)
        {
            try
            {
                string db_check_cache = Check_cache(setID);
                string setName = "";
                if (db_check_cache == "SetIsInCache")
                {

                    string connectionString = @"Data Source=" + DataSource + ";Initial Catalog=" + InitialCatalog + ";User ID=" + DBUser + ";Password=" + DBPassword+ ";MultipleActiveResultSets = true";
                    SqlConnection setNameConnection = new SqlConnection(connectionString);
                    setNameConnection.Open();
                    String sql = "SELECT NAME FROM [dbo].Sets where ID=" + setID + "";
                    SqlCommand command = new SqlCommand(sql, setNameConnection);
                    SqlDataReader setNameReader = command.ExecuteReader();
                    if (setNameReader.HasRows)
                    {
                        while (setNameReader.Read())
                        {
                            setName = Convert.ToString(setNameReader["name"]);
                        }

                    }
                    else
                    {
                        setName = "No value was found";
                    }
                    setNameReader.Close();
                    setNameConnection.Close();
                    return setName;
                }
                else
                {
                    return db_check_cache;
                }                
            } catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item Thumbnail
        public static string GetSetThumbnail(string setID)
        {
            try
            {
                string db_check_cache = Check_cache(setID);
                string setThumbnail = "";
                if (db_check_cache == "SetIsInCache")
                {
                    try
                    {
                        string connectionString = @"Data Source=" + DataSource + ";Initial Catalog=" + InitialCatalog + ";User ID=" + DBUser + ";Password=" + DBPassword + ";MultipleActiveResultSets = true";
                        SqlConnection connection = new SqlConnection(connectionString);
                        connection.Open();
                        String sql = "SELECT thumbnail_url FROM [dbo].Sets where ID=" + setID + "";
                        SqlCommand command = new SqlCommand(sql, connection);
                        SqlDataReader reader = command.ExecuteReader();
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                setThumbnail = Convert.ToString(reader["thumbnail_url"]);
                            }

                        }
                        else
                        {
                            setThumbnail = "No value was found";
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        return ex.Message.ToString();
                    }
                }
                return setThumbnail;
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item image
        public static string GetSetImage(string setID)
        {
            try
            {
                string db_check_cache = Check_cache(setID);
                string setImage = "";
                if (db_check_cache == "SetIsInCache")
                {
                    try
                    {
                        string connectionString = @"Data Source=" + DataSource + ";Initial Catalog=" + InitialCatalog + ";User ID=" + DBUser + ";Password=" + DBPassword + ";MultipleActiveResultSets = true";
                        SqlConnection connection = new SqlConnection(connectionString);
                        connection.Open();
                        String sql = "SELECT imageURL FROM [dbo].Sets where ID=" + setID + "";
                        SqlCommand command = new SqlCommand(sql, connection);
                        SqlDataReader reader = command.ExecuteReader();
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                setImage = Convert.ToString(reader["imageURL"]);
                            }

                        }
                        else
                        {
                            setImage = "No value was found";
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        return ex.Message.ToString();
                    }
                }
                return setImage;
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }


        //This function will return the item release year
        public static string GetSetYear(string setID)
        {
            try
            {
                string db_check_cache = Check_cache(setID);
                string setYear = "";
                if (db_check_cache == "SetIsInCache")
                {
                    try
                    {
                        string connectionString = @"Data Source=" + DataSource + ";Initial Catalog=" + InitialCatalog + ";User ID=" + DBUser + ";Password=" + DBPassword + ";MultipleActiveResultSets = true";
                        SqlConnection connection = new SqlConnection(connectionString);
                        connection.Open();
                        String sql = "SELECT year_released FROM [dbo].Sets where ID=" + setID + "";
                        SqlCommand command = new SqlCommand(sql, connection);
                        SqlDataReader reader = command.ExecuteReader();
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                setYear = Convert.ToString(reader["year_released"]);
                            }

                        }
                        else
                        {
                            setYear = "No value was found";
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        return ex.Message.ToString();
                    }
                }
                return setYear;
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        //This function will return the item average price. Note this code is designed to pull the average price for a new set. Some older sets may not have an entry in BrinkLink so you may want to change this method
        public static string GetSetPrice(string setID)
        {
            string setInformation = GetSetInformation(brickLinkURL + setID + "-1", "price");
            var obj = JObject.Parse(setInformation);
            if (!obj.ToString().Contains("TIMESTAMP"))
            {
                string setValue = (string)obj["data"]["avg_price"];
                if (setValue != null)
                {
                    return setValue;
                }
                else
                {
                    return "no results";
                }
            }
            else
            {
                // Added this sleep function as BrickLink API expect a 0.5 second delay between different call 
                Thread.Sleep(500);
                return GetSetPrice(setID);
            }
        }

        //This function will return the item category ID. 
        //It will compare it to the bricklinkcategorylist.xml file (which is a dump of the /categories API call from bricklink) and return the category name
        public static string GetSetCategory(string setID)
        {
            string setInformation = GetSetInformation(brickLinkURL + setID + "-1", "info");
            var obj = JObject.Parse(setInformation);
            if (!obj.ToString().Contains("TIMESTAMP"))
            {
                string category_id = (string)obj["data"]["category_id"];
                if (category_id != null)
                {
                    XmlReader xmlFile;
                    System.Reflection.Assembly a = System.Reflection.Assembly.GetExecutingAssembly();
                    xmlFile = XmlReader.Create(a.GetManifestResourceStream("BrickLink.bricklinkcategorylist.xml"), new XmlReaderSettings());
                    DataSet ds = new DataSet();
                    DataView dv;
                    ds.ReadXml(xmlFile);

                    dv = new DataView(ds.Tables[0])
                    {
                        Sort = "category_id"
                    };

                    int index = dv.Find(category_id);

                    if (index == -1)
                    {
                        return "no results";
                    }
                    else
                    {
                        return dv[index]["category_name"].ToString();
                    }
                }
                else
                {
                    return "no results";
                }
            }
            else
            {
                // Added this sleep function as BrickLink API expect a 0.5 second delay between different call 
                Thread.Sleep(500);
                return GetSetCategory(setID);
            }
        }               

        public static string Check_cache(string setID)
        // The goal for this new funtion is to streamline the cache usage of the solution. 
        // Instead of making an API call for each field in the excell file,
        // this code will first check if the the set is in the cache and if not, populate the cache for the calling function
        {
            try
            {
                string connectionString = @"Data Source=" + DataSource + ";Initial Catalog=" + InitialCatalog + ";User ID=" + DBUser + ";Password=" + DBPassword + ";MultipleActiveResultSets = true";
                SqlConnection setCacheConnection = new SqlConnection(connectionString);
                setCacheConnection.Open();
                String sql = "SELECT * FROM [dbo].Sets where ID=" + setID;
                SqlCommand command = new SqlCommand(sql, setCacheConnection);
                SqlDataReader cacheReader = command.ExecuteReader();

                DateTime today = DateTime.Now;
                today.ToString("yyyy-MM-dd");
                string SetURL = brickLinkURL + setID + "-1";

                if (!(cacheReader.HasRows))
                {
                    string setNameJSON = GetSetInformation(SetURL, "info");
                    var obj = JObject.Parse(setNameJSON);
                    if (!obj.ToString().Contains("TIMESTAMP"))
                    {
                        string setName = (string)obj["data"]["name"];
                        if (!(setName == null))
                        {
                            // get set information
                            string setType = (string)obj["data"]["type"] != null ? (string)obj["data"]["type"] : "N/A";
                            string setImageURL = (string)obj["data"]["image_url"] != null ? "https:" + (string)obj["data"]["image_url"] : "N/A";
                            string setThumbnailURL = (string)obj["data"]["thumbnail_url"] != null ? "https:" + (string)obj["data"]["thumbnail_url"] : "N/A";
                            string setYear_released = (string)obj["data"]["year_released"] != null ? (string)obj["data"]["year_released"] : "N/A";

                            // Get set price
                            string setPrice = GetSetPrice(setID);

                            // get set category
                            string category_id = GetSetCategory(setID);

                            string insertsql = "INSERT INTO dbo.Sets (ID,name,type,categoryID,imageURL,thumbnail_url,year_released,avg_price,date_updated) VALUES (@ID,@name,@type,@categoryID,@imageURL,@thumbnail_url,@year_released,@avg_price,@date_updated)";
                            SqlCommand insertCommand = new SqlCommand(insertsql, setCacheConnection);
                            insertCommand.Parameters.AddWithValue("@ID", setID);
                            insertCommand.Parameters.AddWithValue("@name", HttpUtility.HtmlDecode(setName));
                            insertCommand.Parameters.AddWithValue("@type", setType);
                            insertCommand.Parameters.AddWithValue("@categoryID", category_id);
                            insertCommand.Parameters.AddWithValue("@imageURL", setImageURL);
                            insertCommand.Parameters.AddWithValue("@thumbnail_url", setThumbnailURL);
                            insertCommand.Parameters.AddWithValue("@year_released", setYear_released);
                            insertCommand.Parameters.AddWithValue("@avg_price", setPrice);
                            insertCommand.Parameters.AddWithValue("@date_updated", today);
                            string debugSQL = insertCommand.ToString();
                            int result = insertCommand.ExecuteNonQuery();

                            // check for errors
                            if (result < 0)
                            {
                                return "There was an error inserting the values to the DB";
                            }
                            setCacheConnection.Close();
                        }
                        return "SetIsInCache";
                    }
                    else
                    {
                        // Added this sleep function as BrickLink API expect a 0.5 second delay between different call 
                        Thread.Sleep(500);
                        return Check_cache(setID);
                    }
                }
                else
                {
                    while (cacheReader.Read())
                    {
                        string dbDate = cacheReader["date_updated"].ToString();
                        DateTime dbDateTime = DateTime.Parse(dbDate);
                        TimeSpan diffOfDays = today - dbDateTime;

                        if (diffOfDays.TotalDays > 30)
                        {
                            string setInformation = GetSetInformation(brickLinkURL + setID + "-1", "info");
                            var obj = JObject.Parse(setInformation);
                            if (!obj.ToString().Contains("TIMESTAMP"))
                            {
                                
                                string setName = (string)obj["data"]["name"];
                                
                                if (!(setName == null))
                                {
                                    // get set information
                                    string setType = (string)obj["data"]["type"] != null ? (string)obj["data"]["type"] : "N/A";
                                    string setImageURL = (string)obj["data"]["image_url"] != null ? "https:" + (string)obj["data"]["image_url"] : "N/A";
                                    string setThumbnailURL = (string)obj["data"]["thumbnail_url"] != null ? "https:" + (string)obj["data"]["thumbnail_url"] : "N/A";
                                    string setYear_released = (string)obj["data"]["year_released"] != null ? (string)obj["data"]["year_released"] : "N/A";

                                    // Get set price
                                    string setPrice = GetSetPrice(setID);

                                    // get set category
                                    string category_id = GetSetCategory(setID);

                                    string updatesql = "update dbo.Sets set name=@name,type=@type,categoryID=@category_id,imageURL=@imageURL,thumbnail_url=@thumbnail_url,year_released=@year_released,avg_price=@avg_price,date_updated=@date_updated where ID=@ID";
                                    SqlCommand updateCommand = new SqlCommand(updatesql, setCacheConnection);
                                    updateCommand.Parameters.AddWithValue("@ID", setID);
                                    updateCommand.Parameters.AddWithValue("@name", HttpUtility.HtmlDecode(setName));
                                    updateCommand.Parameters.AddWithValue("@type", setType);
                                    updateCommand.Parameters.AddWithValue("@category_id", category_id);
                                    updateCommand.Parameters.AddWithValue("@imageURL", setImageURL);
                                    updateCommand.Parameters.AddWithValue("@thumbnail_url", setThumbnailURL);
                                    updateCommand.Parameters.AddWithValue("@year_released", setYear_released);
                                    updateCommand.Parameters.AddWithValue("@avg_price", setPrice);
                                    updateCommand.Parameters.AddWithValue("@date_updated", today);
                                    int result = updateCommand.ExecuteNonQuery();

                                    // check for errors
                                    if (result < 0)
                                    {
                                        return "There was an error inserting the values to the DB";
                                    }                                    
                                }
                            }
                            else
                            {
                                // Added this sleep function as BrickLink API expect a 0.5 second delay between different call 
                                Thread.Sleep(500);
                                return Check_cache(setID);
                            }
                        }
                        else
                        {
                            return "SetIsInCache";                            
                        }
                    }                    
                    return "SetIsInCache";                    
                }                
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }
        public static void Main()
        { }
    }
}
