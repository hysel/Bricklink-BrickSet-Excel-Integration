// ***************************************************************************************
// * BrickLink/BrickeSet Excel function integration 
// * Version 3.1 3/23/2025
// * Itamar Budin brickmindz@gmail.com
// * Using code samples from multiple resources (see internal comments for reference) 
// ***************************************************************************************
// This solution is using the Excel-DNA plug-in. If you would like more details, please look at the ExcelDna.AddIn.md file
// This version of the tool which includes
//  * Introduced new integration with BrickSet 
//  * Major code optimization to reduce code duplication (I know I can do better :)
//      
// Provider HTTPS connections use TLS 1.2 on the legacy .NET Framework build.
// I am not a developer, but I know how to write basic code, so please excuse any lousy code writing :)

using ExcelDna.Integration;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Security.Authentication;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace BrickLinkBrickSet
{
    public static class BricklinkBrickSetExcelIntegration
    {
        // Configure the provider credentials in Windows Credential Manager.
        // For more information see:
        //  BrickLink: https://www.bricklink.com/v3/api.page
        //  BrickSet: https://brickset.com/api/v3.asmx

        // BrickLink 
        private const string CredentialPrefix = "BrickLinkBrickSet/Legacy/";
        const string brickLinkSetURL = "https://api.bricklink.com/api/store/v3/items/set/";             // BrickLink API Set URL
        const string brickLinkGearURL = "https://api.bricklink.com/api/store/v3/items/gear/";           // BrickLink API Gear URL
        const string brickLinkMiniFigURL = "https://api.bricklink.com/api/store/v3/items/minifig/";     // BrickLink API Minifig URL
        const string brickLinkPartURL = "https://api.bricklink.com/api/store/v3/items/part/";           // BrickLink API Part URL
        const string brickLinkBooksURL = "https://api.bricklink.com/api/store/v3/items/book/";          // BrickLink API Book URL
        const string brickLinkCategoryURL = "https://api.bricklink.com/api/store/v3/categories/";       // BrickLink API Book URL        

        // BrickSet
        const string brickSetSOAPUrl = "https://brickset.com/api/v3.asmx";      // Brickset URL

        public static string tokenx = string.Empty;

        // DB Detailes (for cache)
        const string dbNameAttribute = "name";                  // The DB column that holds the set name        
        const string dbTypeAttribute = "type";                  // The DB column that holds the set name        
        const string dbCategoryIDAttribute = "categoryID";      // The DB column that holds the set category  
        const string dbImageURLAttribute = "image_url";          // The DB column that holds the set image URL
        const string dbThumbnailURLAttribute = "thumbnail_url"; // The DB column that holds the set thumbnail URL 
        const string dbYearAttribute = "year_released";         // The DB column that holds the set release year  
        const string dbAvgPriceAttribute = "avg_price";         // The DB column that holds the set average price year (BrickLink only)
        const string dbPartNumberAttribute = "partnum";         // The DB column that holds the set part number     
        const string dbNumOfMinifigsAttribute = "minifignum";   // The DB column that holds the set number of minifig (Bricklink only)    
        const string dbUPCAttribute = "UPC";                    // The DB column that holds the set UPC    
        const string dbDescriptoinAttribute = "description";    // The DB column that holds the set description    
        const string dbOrgPriceAttribute = "original_price";    // The DB column that holds the set original price   
        const string dbSetMinifiguresAttribute = "minifigset";  // The DB columne that holds the set original price   

        [StructLayout(LayoutKind.Sequential)]
        private struct WindowsCredential
        {
            public uint Flags;
            public uint Type;
            public IntPtr TargetName;
            public IntPtr Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public uint CredentialBlobSize;
            public IntPtr CredentialBlob;
            public uint Persist;
            public uint AttributeCount;
            public IntPtr Attributes;
            public IntPtr TargetAlias;
            public IntPtr UserName;
        }

        [DllImport("Advapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CredReadW(string target, uint type, uint flags, out IntPtr credential);

        [DllImport("Advapi32.dll", ExactSpelling = true)]
        private static extern void CredFree(IntPtr credential);

        private static string ReadCredential(string name)
        {
            const uint genericCredential = 1;
            const int credentialNotFound = 1168;
            string target = CredentialPrefix + name;
            if (!CredReadW(target, genericCredential, 0, out IntPtr entry))
            {
                int error = Marshal.GetLastWin32Error();
                throw new LookupException(error == credentialNotFound
                    ? "Missing Windows Credential Manager entry: " + target
                    : $"Could not read Windows Credential Manager entry {target} (Windows error {error}).");
            }

            try
            {
                var credential = Marshal.PtrToStructure<WindowsCredential>(entry);
                return DecodeCredential(credential.CredentialBlob, credential.CredentialBlobSize, target);
            }
            finally
            {
                CredFree(entry);
            }
        }

        private static string DecodeCredential(IntPtr blob, uint byteLength, string target)
        {
            const uint maxCredentialBlobBytes = 2560;
            if (blob == IntPtr.Zero || byteLength == 0 || byteLength > maxCredentialBlobBytes || byteLength % 2 != 0)
            {
                throw new LookupException("Incomplete or invalid Windows Credential Manager entry: " + target);
            }

            string password = Marshal.PtrToStringUni(blob, (int)byteLength / 2).TrimEnd('\0');
            if (string.IsNullOrWhiteSpace(password))
            {
                throw new LookupException("Empty password in Windows Credential Manager entry: " + target);
            }

            return password;
        }

        private static string GetCacheConnectionString()
        {
            return new SqlConnectionStringBuilder
            {
                DataSource = ReadCredential("SqlServer"),
                InitialCatalog = ReadCredential("SqlDatabase"),
                UserID = ReadCredential("SqlUser"),
                Password = ReadCredential("SqlPassword")
            }.ConnectionString;
        }


        private static readonly ConcurrentDictionary<string, Lazy<string>> CacheRefreshes =
            new(StringComparer.Ordinal);
        private static readonly ConcurrentDictionary<string, Lazy<JObject>> BrickSetLookups =
            new(StringComparer.OrdinalIgnoreCase);
        private static readonly SemaphoreSlim ProviderSlots = new(4);
        private static readonly Lazy<HttpClient> ProviderHttpClient = new(() => new HttpClient(CreateProviderHttpHandler())
        {
            Timeout = TimeSpan.FromSeconds(10)
        });
        private const string NotAvailable = "N/A";
        private const string SetNotFound = "No Set Information has been found";
        private const string LatestCacheRowOrder = "ORDER BY [date_updated] DESC, [ID] COLLATE Latin1_General_100_BIN2";

        private sealed class LookupException : InvalidOperationException
        {
            public LookupException(string message) : base(message) { }
        }

        private static string GetLookupError(Exception error)
        {
            if (error is LookupException)
            {
                return error.Message;
            }

            return error is SqlException sqlError
                ? $"SQL cache operation failed (SQL error {sqlError.Number})."
                : "Legacy lookup failed.";
        }

        private static string ValidateSetId(string setID)
        {
            string id = setID?.Trim();
            if (string.IsNullOrEmpty(id))
            {
                throw new LookupException("Enter a set ID.");
            }
            return id;
        }

        private static string ReadSetInformationFromDB(string connectionString, string setID, string column)
        {
            if (column is not (dbNameAttribute or dbTypeAttribute or dbCategoryIDAttribute or
                dbImageURLAttribute or dbThumbnailURLAttribute or dbYearAttribute or dbAvgPriceAttribute or
                dbPartNumberAttribute or dbNumOfMinifigsAttribute or dbUPCAttribute or
                dbDescriptoinAttribute or dbOrgPriceAttribute or dbSetMinifiguresAttribute))
            {
                throw new ArgumentOutOfRangeException(nameof(column));
            }

            using var connection = new SqlConnection(connectionString);
            connection.Open();
            using var command = new SqlCommand(
                "SELECT TOP (1) [" + column + "] FROM [dbo].[Sets] WHERE [ID] = @setID " + LatestCacheRowOrder,
                connection);
            command.Parameters.Add("@setID", SqlDbType.NVarChar, -1).Value = setID;
            using var reader = command.ExecuteReader();
            if (!reader.Read())
            {
                return NotAvailable;
            }

            string value = Convert.ToString(reader.GetValue(0));
            if (column is dbImageURLAttribute or dbThumbnailURLAttribute)
            {
                return NormalizeImageUrl(value);
            }

            // Older fills may have repeated an ID with its last quantity. Refresh those rows from the provider.
            return column == dbSetMinifiguresAttribute && HasDuplicateMinifigureIds(value) ? NotAvailable : value;
        }

        private static bool HasDuplicateMinifigureIds(string value)
        {
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string entry in value.Split(new[] { ", " }, StringSplitOptions.None))
            {
                int quantityStart = entry.LastIndexOf(" (", StringComparison.Ordinal);
                if (quantityStart > 0 && entry.EndsWith(")", StringComparison.Ordinal) &&
                    !ids.Add(entry.Substring(0, quantityStart)))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool CacheFieldIsMissing(string value) =>
            string.IsNullOrWhiteSpace(value) || value is NotAvailable or "no results";

        private static string NormalizeImageUrl(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return NotAvailable;
            }

            value = value.Trim();
            if (value.StartsWith("//", StringComparison.Ordinal))
            {
                value = "https:" + value;
            }

            return Uri.TryCreate(value, UriKind.Absolute, out Uri uri) &&
                (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp) &&
                !string.IsNullOrEmpty(uri.Host)
                ? value
                : NotAvailable;
        }

        private static T WithProviderSlot<T>(Func<T> operation)
        {
            if (!ProviderSlots.Wait(TimeSpan.FromSeconds(30)))
            {
                throw new LookupException("All provider request slots are busy. Recalculate to retry.");
            }
            try { return operation(); }
            finally { ProviderSlots.Release(); }
        }

        private static T RunSharedLookup<T>(ConcurrentDictionary<string, Lazy<T>> pending, string key, Func<T> lookup)
        {
            var candidate = new Lazy<T>(lookup, LazyThreadSafetyMode.ExecutionAndPublication);
            var active = pending.GetOrAdd(key, candidate);
            try { return active.Value; }
            finally
            {
                if (ReferenceEquals(candidate, active))
                {
                    pending.TryRemove(key, out _);
                }
            }
        }

        private static readonly string[] BrickLinkItemUrls =
        {
            brickLinkSetURL, brickLinkGearURL, brickLinkMiniFigURL, brickLinkPartURL, brickLinkBooksURL
        };

        private static string EscapeOAuth(string value) => Uri.EscapeDataString(value)
            .Replace("!", "%21").Replace("*", "%2A").Replace("'", "%27").Replace("(", "%28").Replace(")", "%29");

        private static HttpClientHandler CreateProviderHttpHandler() => new()
        {
            SslProtocols = SslProtocols.Tls12,
            AllowAutoRedirect = false,
            UseCookies = false,
            MaxConnectionsPerServer = 4
        };

        private static HttpResponseMessage SendProviderRequest(HttpRequestMessage request, string service)
        {
            try { return ProviderHttpClient.Value.SendAsync(request).GetAwaiter().GetResult(); }
            catch (OperationCanceledException)
            {
                throw new LookupException($"{service} connection timed out.");
            }
            catch (HttpRequestException error)
            {
                var status = (error.InnerException as WebException)?.Status;
                throw new LookupException(status.HasValue
                    ? $"{service} connection failed ({status.Value})."
                    : $"{service} connection failed.");
            }
        }

        private static (int Code, JObject Body) RequestBrickLink(string url)
        {
            string consumerKey = ReadCredential("BrickLinkConsumerKey");
            string consumerSecret = ReadCredential("BrickLinkConsumerSecret");
            string token = ReadCredential("BrickLinkToken");
            string tokenSecret = ReadCredential("BrickLinkTokenSecret");
            string timestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds().ToString(CultureInfo.InvariantCulture);
            string nonce = Guid.NewGuid().ToString("N");
            var oauth = new[]
            {
                ("oauth_consumer_key", consumerKey),
                ("oauth_nonce", nonce),
                ("oauth_signature_method", "HMAC-SHA1"),
                ("oauth_timestamp", timestamp),
                ("oauth_token", token),
                ("oauth_version", "1.0")
            };
            string parameters = string.Join("&", oauth.Select(pair => pair.Item1 + "=" + EscapeOAuth(pair.Item2)));
            string signatureBase = "GET&" + EscapeOAuth(url) + "&" + EscapeOAuth(parameters);
            byte[] key = Encoding.ASCII.GetBytes(EscapeOAuth(consumerSecret) + "&" + EscapeOAuth(tokenSecret));
            string signature;
            try
            {
                using var hmac = new HMACSHA1(key);
                signature = Convert.ToBase64String(hmac.ComputeHash(Encoding.ASCII.GetBytes(signatureBase)));
            }
            finally { Array.Clear(key, 0, key.Length); }

            using var request = new HttpRequestMessage(HttpMethod.Get, url);
            string header = string.Join(",", oauth.Select(pair => pair.Item1 + "=\"" + EscapeOAuth(pair.Item2) + "\""));
            request.Headers.Authorization = new AuthenticationHeaderValue(
                "OAuth", header + ",oauth_signature=\"" + EscapeOAuth(signature) + "\"");
            using var response = SendProviderRequest(request, "BrickLink");
            string content = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            return ReadBrickLinkResponse((int)response.StatusCode, content);
        }

        private static (int Code, JObject Body) ReadBrickLinkResponse(int status, string content)
        {
            bool httpSuccess = status >= 200 && status < 300;
            try
            {
                var body = JObject.Parse(content);
                var meta = body["meta"] as JObject;
                int code = httpSuccess && int.TryParse((string)meta?["code"], out int apiCode) ? apiCode : status;
                return (code, body);
            }
            catch (Newtonsoft.Json.JsonException)
            {
                if (httpSuccess)
                {
                    throw new LookupException("BrickLink returned an invalid response.");
                }
                return (status, null);
            }
        }

        private static JToken GetBrickLinkData(string url, bool allowMissing = false) =>
            GetBrickLinkData(() => RequestBrickLink(url), allowMissing);

        private static T GetOptionalBrickLinkData<T>(string url) where T : JToken
        {
            JToken data = GetBrickLinkData(url, allowMissing: true);
            if (data == null)
            {
                return null;
            }
            return data as T ?? throw new LookupException("BrickLink returned an unexpected data format.");
        }

        private static JToken GetBrickLinkData(Func<(int Code, JObject Body)> request, bool allowMissing)
        {
            const int maxAttempts = 3;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                var (code, body) = request();
                var data = body?["data"];
                if (code >= 200 && code < 300 && data != null && data.Type != JTokenType.Null)
                {
                    return data;
                }

                var meta = body?["meta"] as JObject;
                string error = (string)meta?["message"] + " " + (string)meta?["description"];
                if (error.IndexOf("TIMESTAMP", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (attempt == maxAttempts)
                    {
                        throw new LookupException("BrickLink rejected the request timestamp after three attempts.");
                    }
                    Thread.Sleep(500);
                    continue;
                }

                if (allowMissing && code is (400 or 404))
                {
                    return null;
                }

                throw new LookupException(code >= 200 && code < 300
                    ? "BrickLink did not return the requested data."
                    : $"BrickLink request failed (status {code}).");
            }
            throw new LookupException("BrickLink request failed.");
        }

        private static (JObject Item, string ItemUrl) FindBrickLinkItem(string setID)
        {
            string escaped = Uri.EscapeDataString(setID);
            string url = brickLinkSetURL + escaped + "-1";
            if (GetOptionalBrickLinkData<JObject>(url) is JObject variant)
            {
                return (variant, url);
            }

            foreach (string prefix in BrickLinkItemUrls)
            {
                url = prefix + escaped;
                if (GetOptionalBrickLinkData<JObject>(url) is JObject item)
                {
                    return (item, url);
                }
            }
            return (null, null);
        }

        private static (int Count, string Collection) SummarizeMinifigures(JArray subsets)
        {
            var quantities = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var order = new List<string>();
            int total = 0;
            if (subsets != null)
            {
                foreach (var subset in subsets.OfType<JObject>())
                {
                    if (subset["entries"] is not JArray entries)
                    {
                        continue;
                    }

                    foreach (var entry in entries.OfType<JObject>())
                    {
                        var item = entry["item"] as JObject;
                        string id = (string)item?["no"];
                        if (!string.Equals((string)item?["type"], "MINIFIG", StringComparison.OrdinalIgnoreCase) ||
                            string.IsNullOrWhiteSpace(id) || !int.TryParse((string)entry["quantity"], out int quantity) ||
                            quantity <= 0)
                        {
                            continue;
                        }

                        if (!quantities.TryGetValue(id, out int previous))
                        {
                            order.Add(id);
                        }
                        quantities[id] = checked(previous + quantity);
                        total = checked(total + quantity);
                    }
                }
            }
            return (total, string.Join(", ", order.Select(id => $"{id} ({quantities[id]})")));
        }

        //This function will return the item name
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetNameFromBrickLink(string setID)
        {
            return ReadCachedSetField(setID, dbNameAttribute);
        }

        // Return the total number of minifigures for the set.
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetMiniFigNumberFromBrickLink(string setID)
        {
            return ReadCachedSetField(setID, dbNumOfMinifigsAttribute);
        }

        // Return the minifigure IDs and their quantities.
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetMiniFigCollectionFromBrickLink(string setID)
        {
            return ReadCachedSetField(setID, dbSetMinifiguresAttribute);
        }

        //This function will return the item Thumbnail
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetThumbnailFromBrickLink(string setID)
        {
            return ReadCachedSetField(setID, dbThumbnailURLAttribute);
        }

        //This function will return the item image
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetImageFromBrickLink(string setID)
        {
            return ReadCachedSetField(setID, dbImageURLAttribute);
        }

        //This function will return the item release year
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetReleaseYearFBrickLink(string setID)
        {
            return ReadCachedSetField(setID, dbYearAttribute);
        }

        //This function will return the item type (Set, gear, etc.)
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetTypeFromBrickLink(string setID)
        {
            return ReadCachedSetField(setID, dbTypeAttribute);
        }

        //This function will return the item average price. Note this code is designed to pull the average price for a new set. Some older sets may not have an entry in BrinkLink so you may want to change this method
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetAvgPriceFromBrickLink(string setID)
        {
            return ReadCachedSetField(setID, dbAvgPriceAttribute);
        }

        //This function will return the item category ID. 
        //It will compare it to the /categories API call from bricklink, and return the category name
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetCategoryFromBrickLink(string setID)
        {
            return ReadCachedSetField(setID, dbCategoryIDAttribute);
        }

        public static string updateSetInCache(string setID, string dbColumn)
        {
            try
            {
                string id = ValidateSetId(setID);
                string connectionString = GetCacheConnectionString();
                return RunSharedLookup(CacheRefreshes, id, () => PopulateSetCache(connectionString, id));
            }
            catch (Exception error) { return GetLookupError(error); }
        }

        private static string PopulateSetCache(string connectionString, string setID)
        {
            var item = WithProviderSlot(() => FetchSetCacheData(setID));
            if (item == null)
            {
                return SetNotFound;
            }

            for (int attempt = 1; ; attempt++)
            {
                try { return WriteSetCache(connectionString, setID, item); }
                catch (SqlException error) when (error.Number == 1205 && attempt < 3)
                {
                    Thread.Sleep(50 * attempt);
                }
            }
        }

        private static Dictionary<string, object> FetchSetCacheData(string setID)
        {
            var (item, itemUrl) = FindBrickLinkItem(setID);
            string name = WebUtility.HtmlDecode((string)item?["name"]);
            if (string.IsNullOrWhiteSpace(name) || name == NotAvailable)
            {
                return null;
            }

            var prices = GetOptionalBrickLinkData<JObject>(itemUrl + "/price");
            string categoryId = (string)item["category_id"];
            var category = string.IsNullOrWhiteSpace(categoryId) ? null :
                GetOptionalBrickLinkData<JObject>(brickLinkCategoryURL + Uri.EscapeDataString(categoryId));
            var subsets = GetOptionalBrickLinkData<JArray>(itemUrl + "/subsets");
            var (minifigureCount, minifigures) = SummarizeMinifigures(subsets);

            JObject brickSet = GetBrickSetItem(setID);
            var barcode = brickSet?["barcode"] as JObject;
            var image = brickSet?["image"] as JObject;
            var details = brickSet?["extendedData"] as JObject;
            string upc = (string)barcode?["UPC"];
            string description = WebUtility.HtmlDecode((string)details?["description"]);
            string imageUrl = NormalizeImageUrl((string)item["image_url"]);
            string thumbnail = NormalizeImageUrl((string)item["thumbnail_url"]);
            return new Dictionary<string, object>
            {
                [dbNameAttribute] = name,
                [dbTypeAttribute] = (string)item["type"],
                [dbCategoryIDAttribute] = WebUtility.HtmlDecode((string)category?["category_name"]),
                [dbImageURLAttribute] = imageUrl == NotAvailable ? NormalizeImageUrl((string)image?["imageURL"]) : imageUrl,
                [dbThumbnailURLAttribute] = thumbnail == NotAvailable ? NormalizeImageUrl((string)image?["thumbnailURL"]) : thumbnail,
                [dbYearAttribute] = (string)item["year_released"],
                [dbAvgPriceAttribute] = (string)prices?[dbAvgPriceAttribute],
                ["date_updated"] = DateTime.Now,
                [dbPartNumberAttribute] = int.TryParse((string)brickSet?["pieces"], out int pieces) ? (int?)pieces : null,
                [dbNumOfMinifigsAttribute] = minifigureCount,
                [dbUPCAttribute] = CacheFieldIsMissing(upc) ? (string)barcode?["EAN"] : upc,
                [dbDescriptoinAttribute] = description?.Replace("<p>", "").Replace("</p>", ""),
                [dbOrgPriceAttribute] = ExtractBrickSetRetailPrice(brickSet, "US"),
                [dbSetMinifiguresAttribute] = minifigures
            };
        }

        private static string WriteSetCache(string connectionString, string setID, Dictionary<string, object> fields)
        {
            const string sql = @"
;WITH LatestSet AS (
    SELECT TOP (1) * FROM [dbo].[Sets] WITH (UPDLOCK, HOLDLOCK)
    WHERE [ID] = @setID
    " + LatestCacheRowOrder + @"
)
UPDATE LatestSet SET
    [name] = @name, [type] = @type, [categoryID] = @categoryID, [image_url] = @image_url,
    [thumbnail_url] = @thumbnail_url, [year_released] = @year_released, [avg_price] = @avg_price,
    [date_updated] = @date_updated, [partnum] = @partnum, [minifignum] = @minifignum,
    [UPC] = @UPC, [description] = @description, [original_price] = @original_price, [minifigset] = @minifigset;

IF @@ROWCOUNT = 0
BEGIN
    INSERT INTO [dbo].[Sets]
        ([ID], [name], [type], [categoryID], [image_url], [thumbnail_url], [year_released], [avg_price],
         [date_updated], [partnum], [minifignum], [UPC], [description], [original_price], [minifigset])
    VALUES
        (@setID, @name, @type, @categoryID, @image_url, @thumbnail_url, @year_released, @avg_price,
         @date_updated, @partnum, @minifignum, @UPC, @description, @original_price, @minifigset);
    SELECT CASE WHEN @@ROWCOUNT = 1 THEN 1 ELSE 0 END;
END
ELSE SELECT 2;";
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            using var transaction = connection.BeginTransaction();
            using var command = new SqlCommand(sql, connection, transaction);
            command.Parameters.Add("@setID", SqlDbType.NVarChar, -1).Value = setID;
            foreach (var field in fields)
            {
                var type = field.Key switch
                {
                    "date_updated" => SqlDbType.DateTime,
                    dbPartNumberAttribute or dbNumOfMinifigsAttribute => SqlDbType.Int,
                    _ => SqlDbType.NVarChar
                };
                var parameter = command.Parameters.Add("@" + field.Key, type);
                if (type == SqlDbType.NVarChar)
                {
                    parameter.Size = -1;
                }
                parameter.Value = field.Value ?? DBNull.Value;
            }

            switch (command.ExecuteScalar())
            {
                case 1:
                    transaction.Commit();
                    return "Entry was added to the DB";
                case 2:
                    transaction.Commit();
                    return "DB Record updated";
                default:
                    throw new LookupException("The SQL cache did not save the set.");
            }
        }

        private static JObject GetBrickSetItem(string setID)
        {
            string number = HasSetVariant(setID) ? setID : setID + "-1";
            return RunSharedLookup(BrickSetLookups, number, () => RequestBrickSetItem(number));
        }

        private static JObject RequestBrickSetItem(string number)
        {
            var parameters = new JObject { ["setNumber"] = number, ["extendedData"] = 1 };
            string apiKey = ReadCredential("BrickSetApiKey");
            string userHash = ReadCredential("BrickSetUserHash");
            byte[] form = Encoding.UTF8.GetBytes(
                "apiKey=" + WebUtility.UrlEncode(apiKey) +
                "&userHash=" + WebUtility.UrlEncode(userHash) +
                "&params=" + WebUtility.UrlEncode(parameters.ToString(Newtonsoft.Json.Formatting.None)));
            try
            {
                using var request = new HttpRequestMessage(HttpMethod.Post, brickSetSOAPUrl + "/getSets");
                request.Content = new ByteArrayContent(form);
                request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");
                using var response = SendProviderRequest(request, "BrickSet");
                if (!response.IsSuccessStatusCode)
                {
                    throw new LookupException($"BrickSet request failed (status {(int)response.StatusCode}).");
                }
                return ParseBrickSetItem(response.Content.ReadAsStringAsync().GetAwaiter().GetResult(), number);
            }
            finally { Array.Clear(form, 0, form.Length); }
        }

        private static bool HasSetVariant(string setID)
        {
            int separator = setID.LastIndexOf('-');
            return separator > 0 && int.TryParse(setID.Substring(separator + 1), NumberStyles.None, CultureInfo.InvariantCulture, out _);
        }

        private static JObject ParseBrickSetItem(string response, string setNumber)
        {
            if (string.IsNullOrWhiteSpace(response))
            {
                throw new LookupException("BrickSet returned an empty response.");
            }

            JObject payload;
            try
            {
                using var text = new StringReader(response.TrimStart('\uFEFF', ' ', '\r', '\n', '\t'));
                using var json = new Newtonsoft.Json.JsonTextReader(text);
                payload = JObject.Load(json);
            }
            catch (Newtonsoft.Json.JsonException)
            {
                throw new LookupException("BrickSet returned an invalid response.");
            }

            string status = (string)payload["status"];
            if (!string.Equals(status, "success", StringComparison.OrdinalIgnoreCase))
            {
                string error = (string)payload["message"] ?? string.Empty;
                throw new LookupException(error.IndexOf("limit exceeded", StringComparison.OrdinalIgnoreCase) >= 0
                    ? "BrickSet API limit exceeded"
                    : "BrickSet rejected the lookup. Check the API key and user hash.");
            }

            if (payload["sets"] is not JArray sets)
            {
                if ((int?)payload["matches"] == 0)
                {
                    return null;
                }
                throw new LookupException("BrickSet did not return the expected set results.");
            }

            return sets.OfType<JObject>().FirstOrDefault(item =>
            {
                string number = (string)item["number"];
                string variant = (string)item["numberVariant"];
                return !string.IsNullOrEmpty(number) &&
                    (string.Equals(number, setNumber, StringComparison.OrdinalIgnoreCase) ||
                     (!string.IsNullOrEmpty(variant) &&
                      string.Equals(number + "-" + variant, setNumber, StringComparison.OrdinalIgnoreCase)));
            });
        }

        private static string ExtractBrickSetRetailPrice(JObject item, string country)
        {
            var markets = item?["LEGOCom"] as JObject;
            var market = markets?[country] as JObject;
            return (string)market?["retailPrice"] ?? NotAvailable;
        }

        //This function will return the item name
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetNameFromBrickSet(string setID)
        {
            return ReadCachedSetField(setID, dbNameAttribute, preserveBrickSetLimitMessage: true);
        }

        private static string ReadCachedSetField(string setID, string column, bool preserveBrickSetLimitMessage = false)
        {
            try
            {
                if (preserveBrickSetLimitMessage && setID == "BrickSet API limit exceeded")
                {
                    return setID;
                }

                string id = ValidateSetId(setID);
                string connectionString = GetCacheConnectionString();
                string cached = ReadSetInformationFromDB(connectionString, id, column);
                if (!CacheFieldIsMissing(cached))
                {
                    return cached;
                }

                string refresh = RunSharedLookup(CacheRefreshes, id, () =>
                    CacheFieldIsMissing(ReadSetInformationFromDB(connectionString, id, column))
                        ? PopulateSetCache(connectionString, id)
                        : "Cache already populated");
                if (refresh is not ("Entry was added to the DB" or "DB Record updated" or "Cache already populated"))
                {
                    return refresh;
                }

                cached = ReadSetInformationFromDB(connectionString, id, column);
                return CacheFieldIsMissing(cached) ? NotAvailable : cached;
            }
            catch (Exception error) { return GetLookupError(error); }
        }

        //This function will return the item theme
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetThemeFromBrickSet(string setID)
        {
            return ReadCachedSetField(setID, dbCategoryIDAttribute, preserveBrickSetLimitMessage: true);
        }

        //This function will return the item image URL
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetImageURLFromBrickSet(string setID)
        {
            return ReadCachedSetField(setID, dbImageURLAttribute, preserveBrickSetLimitMessage: true);
        }

        //This function will return the item thumbnail URL 
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetThumbnailURLFromBrickSet(string setID)
        {
            return ReadCachedSetField(setID, dbThumbnailURLAttribute, preserveBrickSetLimitMessage: true);
        }

        //This function will return the item release year 
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetReleaseYearFromBrickSet(string setID)
        {
            return ReadCachedSetField(setID, dbYearAttribute, preserveBrickSetLimitMessage: true);
        }

        //This function will return the item part number
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetPartsNumberFromBrickSet(string setID)
        {
            return ReadCachedSetField(setID, dbPartNumberAttribute, preserveBrickSetLimitMessage: true);
        }

        //This function will return the item UPC 
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetUPCFromBrickSet(string setID)
        {
            return ReadCachedSetField(setID, dbUPCAttribute, preserveBrickSetLimitMessage: true);
        }

        //This function will return the item description
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetDescriptionFromBrickSet(string setID)
        {
            return ReadCachedSetField(setID, dbDescriptoinAttribute, preserveBrickSetLimitMessage: true);
        }

        //This function will return the item original retail price
        [ExcelFunction(IsThreadSafe = true)]
        public static string GetSetOriginalPriceFromBrickSet(string setID, string country)
        {
            if (country is not ("US" or "UK" or "DE" or "CA"))
            {
                return "Wrong country code entered, accetable values are: US, UK, CA and DE";
            }

            if (country == "US")
            {
                return ReadCachedSetField(setID, dbOrgPriceAttribute, preserveBrickSetLimitMessage: true);
            }

            try
            {
                if (setID == "BrickSet API limit exceeded")
                {
                    return setID;
                }
                string id = ValidateSetId(setID);
                return WithProviderSlot(() => ExtractBrickSetRetailPrice(GetBrickSetItem(id), country));
            }
            catch (Exception error) { return GetLookupError(error); }
        }

        public static void Main()
        {
        }
    }
}
