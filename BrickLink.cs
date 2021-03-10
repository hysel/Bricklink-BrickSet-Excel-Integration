// ***************************************************************************************
// * BrickLink Excel function integration 
// * Version 1.0 3/10/2021
// * Itamar Budin lego.c.israel@gmail.com
// * Using code samples from multiple resource (see internal comments for reference) 
// ***************************************************************************************
// This solution is using the Excel-DNA plug-in. For more details, see the ExcelDna.AddIn.md file

// I am not a developer but know how to write basic code so please excuse any bad code writing :)



using Nancy.Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace BrickLink
{
    public static class BricklinkExcelIntegration
    {
        // TODO: In this section, you will need to enter the various secrets and keys that are assigned to you by Bricklink
        // For more information see: https://www.bricklink.com/v2/api/welcome.page

        const string consumerKey = "";
        const string consumerSecret = "";
        const string tokenValue = "";
        const string tokenSecret = "";
        public static string tokenx = "";

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
         * https://api.bricklink.com/api/store/v1/items/set/10030-1
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
            httpWebRequest.KeepAlive = false;
            httpWebRequest.ServicePoint.Expect100Continue = false;

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
            var response = httpWebRequest.GetResponse();
            string characterSet = ((HttpWebResponse)response).CharacterSet;
            var responseEncoding = characterSet == ""
                ? Encoding.UTF8
                : Encoding.GetEncoding(characterSet ?? "utf-8");
            var responsestream = response.GetResponseStream();
            if (responsestream == null)
            {
                throw new ArgumentNullException(nameof(characterSet));
            }
            using (responsestream)
            {
                var reader = new StreamReader(responsestream, responseEncoding);
                string result = reader.ReadToEnd();
                Console.WriteLine(@"result: " + result);


                Regex tid = new Regex("oauth_token=(.*?)&");
                Match match = tid.Match(result);
                tokenx = match.Groups[1].Value;
                Console.WriteLine(match.Groups[1].Value);
                return result;
            }

        }

        //This function will return the item name
        public static string GetSetName(string name)
        {
            string setInformation = GetSetInformation(GetURL(name), "info");
            var obj = JObject.Parse(setInformation);
            if (!obj.ToString().Contains("TIMESTAMP"))
            {
                string value = (string)obj["data"]["name"];
                return value != null ? HttpUtility.HtmlDecode(value) : "no results";
            }
            else
            {
                // Added this sleep function as BrickLink API expect a 1 second delay between different call 
                Thread.Sleep(500);
                return GetSetName(name);
            }
        }

        //This function will return the item image
        public static string GetSetImage(string name)
        {
            string setInformation = GetSetInformation(GetURL(name), "info");
            var obj = JObject.Parse(setInformation);
            if (!obj.ToString().Contains("TIMESTAMP"))
            {
                string value = (string)obj["data"]["thumbnail_url"];
                if (value != null)
                {
                    return "http:" + value;
                }
                else
                {
                    return "no results";
                }
            }
            else
            {
                // Added this sleep function as BrickLink API expect a 1 second delay between different call 
                Thread.Sleep(500);
                return GetSetImage(name);
            }

        }

        //This function will return the item release year
        public static string GetSetYear(string name)
        {
            string setInformation = GetSetInformation(GetURL(name), "info");
            var obj = JObject.Parse(setInformation);
            if (!obj.ToString().Contains("TIMESTAMP"))
            {
                string value = (string)obj["data"]["year_released"];
                if (value != null)
                {
                    return value;
                }
                else
                {
                    return "no results";
                }
            }
            else
            {
                // Added this sleep function as BrickLink API expect a 1 second delay between different call 
                Thread.Sleep(500);
                return GetSetYear(name);
            }

        }

        //This function will return the item average price. Note this code is designed to pull the average price for a new set. Some older sets may not have an entry in BrinkLink so you may want to change this method
        public static string GetSetPrice(string name)
        {
            string setInformation = GetSetInformation(GetURL(name), "price");
            var obj = JObject.Parse(setInformation);
            if (!obj.ToString().Contains("TIMESTAMP"))
            {
                string value = (string)obj["data"]["avg_price"];
                if (value != null)
                {
                    return value;
                }
                else
                {
                    return "no results";
                }
            }
            else
            {
                // Added this sleep function as BrickLink API expect a 1 second delay between different call 
                Thread.Sleep(500);
                return GetSetPrice(name);
            }
        }

        // This function will build the right API URL based on the set identifier (set#, set code etc.)
        private static string GetURL(string name)
        {
            string url;
            string setTypeValue = GetSetType(name);
            if (setTypeValue == "SET")
            {
                // Some items are categorized as sets even if they don't have an actual number (10030) next to them. this section will help identify them and build the right URL for them
                if (Char.IsLetter(name[0]))
                {
                    url = "https://api.bricklink.com/api/store/v1/items/set/" + name;
                }
                else
                {
                    url = "https://api.bricklink.com/api/store/v1/items/set/" + name + "-1";
                }
            }
            else if (setTypeValue == "GEAR")
            {
                url = "https://api.bricklink.com/api/store/v1/items/gear/" + name;
            }
            else
            {
                url = "https://api.bricklink.com/api/store/v1/items/minifig/" + name;
            }
            return url;
        }

        // This function will return the item type
        // As BrickLink doesn't have an API method to identify an item based on the unique identifier (set # etc.), we will need to use a web client to try to see if we can use BrickLink 
        // URL structure to get to the product page. 
        // Note that this will work using the current BrickLink URL structure and if BrickLink will go and change the pages structure, especially the <Title> tags, this section will need to be updated
        private static string GetSetType(string name)
        {

            WebClient myWebClient = new WebClient();
            string responseBodySet = myWebClient.DownloadString("https://www.bricklink.com/v2/catalog/catalogitem.page?S=" + name);
            string responseBodyGear = myWebClient.DownloadString("https://www.bricklink.com/v2/catalog/catalogitem.page?G=" + name);
            string returnValue = "";

            if ((responseBodySet.Contains("BrickLink Page Not Found") == false)){
                if (responseBodySet.Contains("BrickLink - Oops, Sorry!") == false)
                {
                    returnValue = "SET";
                }            
            }
            else if ((responseBodyGear.Contains("BrickLink Page Not Found") == false))
            {
                if (responseBodyGear.Contains("BrickLink - Oops, Sorry!") == false)
                {
                    returnValue = "GEAR";
                }
            }
            else
            {
                returnValue = "MINI";
            }
            return returnValue;

        }

        public static void Main()
        {

        }

    }

}