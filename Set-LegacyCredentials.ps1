$ErrorActionPreference = 'Stop'

if (-not ('BrickLinkLegacyCredentialStore' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security;

public static class BrickLinkLegacyCredentialStore
{
    private const uint Generic = 1;
    private const int MaxBlobBytes = 2560;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct Credential
    {
        public uint Flags;
        public uint Type;
        public string TargetName;
        public string Comment;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
        public uint BlobSize;
        public IntPtr Blob;
        public uint Persist;
        public uint AttributeCount;
        public IntPtr Attributes;
        public string TargetAlias;
        public string UserName;
    }

    [DllImport("Advapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CredWriteW(ref Credential credential, uint flags);

    [DllImport("Advapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CredReadW(string target, uint type, uint flags, out IntPtr credential);

    [DllImport("Advapi32.dll", ExactSpelling = true)]
    private static extern void CredFree(IntPtr credential);

    public static string GetStatus(string target)
    {
        const int ErrorNotFound = 1168;
        IntPtr entry;
        if (!CredReadW(target, Generic, 0, out entry))
            return Marshal.GetLastWin32Error() == ErrorNotFound ? "Missing" : "Unavailable";

        try
        {
            // Read only the blob length. The menu does not inspect the password bytes.
            int offset = Marshal.OffsetOf(typeof(Credential), "BlobSize").ToInt32();
            int bytes = Marshal.ReadInt32(entry, offset);
            return bytes > 0 && bytes <= MaxBlobBytes && bytes % 2 == 0 ? "Set" : "Incomplete";
        }
        finally { CredFree(entry); }
    }

    public static string Read(string target)
    {
        IntPtr entry;
        if (!CredReadW(target, Generic, 0, out entry))
            throw new InvalidOperationException("Credential entry could not be read.");

        try
        {
            int sizeOffset = Marshal.OffsetOf(typeof(Credential), "BlobSize").ToInt32();
            int blobOffset = Marshal.OffsetOf(typeof(Credential), "Blob").ToInt32();
            int bytes = Marshal.ReadInt32(entry, sizeOffset);
            IntPtr blob = Marshal.ReadIntPtr(entry, blobOffset);
            if (blob == IntPtr.Zero || bytes <= 0 || bytes > MaxBlobBytes || bytes % 2 != 0)
                throw new InvalidOperationException("Credential entry is incomplete.");

            string value = Marshal.PtrToStringUni(blob, bytes / 2).TrimEnd('\0');
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidOperationException("Credential entry is empty.");
            return value;
        }
        finally { CredFree(entry); }
    }

    public static void Save(string target, SecureString value)
    {
        const uint PersistForUserOnThisComputer = 2;
        int bytes = value.Length * sizeof(char);
        if (bytes == 0 || bytes > MaxBlobBytes)
            throw new ArgumentException("Credential value has an invalid length.");

        IntPtr buffer = Marshal.SecureStringToCoTaskMemUnicode(value);
        try
        {
            var entry = new Credential
            {
                Type = Generic,
                TargetName = target,
                Blob = buffer,
                BlobSize = (uint)bytes,
                Persist = PersistForUserOnThisComputer,
                UserName = "BrickLinkBrickSet"
            };
            if (!CredWriteW(ref entry, 0))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        finally
        {
            Marshal.ZeroFreeCoTaskMemUnicode(buffer);
        }
    }
}
'@
}

$entries = [ordered]@{
    BrickLinkConsumerKey    = 'BrickLink consumer key'
    BrickLinkConsumerSecret = 'BrickLink consumer secret'
    BrickLinkToken          = 'BrickLink token'
    BrickLinkTokenSecret    = 'BrickLink token secret'
    BrickSetApiKey         = 'BrickSet API key'
    BrickSetUserHash       = 'BrickSet user hash'
    SqlUser                = 'SQL cache username'
    SqlPassword            = 'SQL cache password'
    SqlServer              = 'SQL Server name or instance'
    SqlDatabase            = 'SQL cache database name'
}

$options = @($entries.GetEnumerator())
$targetPrefix = 'BrickLinkBrickSet/Legacy/'

function ConvertTo-OAuthValue {
    param([string]$Value)
    [Uri]::EscapeDataString($Value).Replace("!", "%21").Replace("'", "%27").Replace("(", "%28").Replace(")", "%29").Replace("*", "%2A")
}

function Test-LegacyBrickLinkConnection {
    param([hashtable]$Values)
    $uri = 'https://api.bricklink.com/api/store/v3/items/set/75192-1'
    $oauth = [ordered]@{
        oauth_consumer_key     = $Values.BrickLinkConsumerKey
        oauth_nonce            = [Guid]::NewGuid().ToString('N')
        oauth_signature_method = 'HMAC-SHA1'
        oauth_timestamp        = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds().ToString([Globalization.CultureInfo]::InvariantCulture)
        oauth_token            = $Values.BrickLinkToken
        oauth_version          = '1.0'
    }
    $parameters = ($oauth.GetEnumerator() | ForEach-Object {
        $_.Key + '=' + (ConvertTo-OAuthValue $_.Value)
    }) -join '&'
    $message = 'GET&' + (ConvertTo-OAuthValue $uri) + '&' + (ConvertTo-OAuthValue $parameters)
    $key = (ConvertTo-OAuthValue $Values.BrickLinkConsumerSecret) + '&' + (ConvertTo-OAuthValue $Values.BrickLinkTokenSecret)
    $keyBytes = [Text.Encoding]::ASCII.GetBytes($key)
    $signer = [Security.Cryptography.HMACSHA1]::new($keyBytes)
    try {
        $oauth.oauth_signature = [Convert]::ToBase64String($signer.ComputeHash([Text.Encoding]::ASCII.GetBytes($message)))
    }
    finally {
        $signer.Dispose()
        [Array]::Clear($keyBytes, 0, $keyBytes.Length)
    }
    $authorization = 'OAuth ' + (($oauth.GetEnumerator() | ForEach-Object {
        '{0}="{1}"' -f $_.Key, (ConvertTo-OAuthValue $_.Value)
    }) -join ',')
    $result = Invoke-RestMethod -Method Get -Uri $uri -Headers @{ Authorization = $authorization } -MaximumRedirection 0 -TimeoutSec 10
    return ($result.meta.code -ge 200 -and $result.meta.code -lt 300 -and $result.data.no -eq '75192-1')
}

function Test-LegacyBrickSetConnection {
    param([hashtable]$Values)
    $body = @{ apiKey = $Values.BrickSetApiKey; userHash = $Values.BrickSetUserHash }
    $result = Invoke-RestMethod -Method Post -Uri 'https://brickset.com/api/v3.asmx/checkUserHash' -Body $body -ContentType 'application/x-www-form-urlencoded' -MaximumRedirection 0 -TimeoutSec 10
    return ($result.status -ceq 'success')
}

function Test-LegacySqlConnection {
    param([hashtable]$Values)
    $settings = [Data.SqlClient.SqlConnectionStringBuilder]::new()
    $settings['Data Source'] = $Values.SqlServer
    $settings['Initial Catalog'] = $Values.SqlDatabase
    $settings['User ID'] = $Values.SqlUser
    $settings['Password'] = $Values.SqlPassword
    $settings['Connect Timeout'] = 10
    $connection = [Data.SqlClient.SqlConnection]::new($settings.ConnectionString)
    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        try {
            $command.CommandText = 'SELECT TOP (0) [ID] FROM [dbo].[Sets]'
            $command.CommandTimeout = 10
            [void]$command.ExecuteNonQuery()
            return $true
        }
        finally { $command.Dispose() }
    }
    finally { $connection.Dispose() }
}

function Show-ConnectionResult {
    param([string]$Service, [string[]]$Entries, [scriptblock]$Test)
    $values = @{}
    try {
        try {
            foreach ($name in $Entries) {
                $values[$name] = [BrickLinkLegacyCredentialStore]::Read($targetPrefix + $name)
            }
        }
        catch {
            Write-Host "$Service - Could not read required Credential Manager entries." -ForegroundColor Yellow
            return
        }

        try { $passed = & $Test $values }
        catch { $passed = $false }
        if ($passed) {
            Write-Host "$Service - Passed" -ForegroundColor Green
        }
        else {
            Write-Host "$Service - Failed. Check the connection settings and service access." -ForegroundColor Yellow
        }
    }
    finally { $values.Clear() }
}

function Show-ConnectionResults {
    Write-Host 'Testing BrickLink, BrickSet, and SQL. This may take a moment.'
    Show-ConnectionResult 'BrickLink' @('BrickLinkConsumerKey', 'BrickLinkConsumerSecret', 'BrickLinkToken', 'BrickLinkTokenSecret') { param($Values) Test-LegacyBrickLinkConnection $Values }
    Show-ConnectionResult 'BrickSet' @('BrickSetApiKey', 'BrickSetUserHash') { param($Values) Test-LegacyBrickSetConnection $Values }
    Show-ConnectionResult 'SQL' @('SqlServer', 'SqlDatabase', 'SqlUser', 'SqlPassword') { param($Values) Test-LegacySqlConnection $Values }
}

:MainMenu while ($true) {
    Write-Host ''
    Write-Host '+------------------------------------------+' -ForegroundColor Cyan
    Write-Host '|      BrickLink + BrickSet for Excel      |' -ForegroundColor Cyan
    Write-Host '|         Legacy credential setup          |' -ForegroundColor Cyan
    Write-Host '+------------------------------------------+' -ForegroundColor Cyan
    for ($i = 0; $i -lt $options.Count; $i++) {
        $entry = $options[$i]
        $status = [BrickLinkLegacyCredentialStore]::GetStatus($targetPrefix + $entry.Key)
        Write-Host ('  {0,2}. {1,-29} ' -f ($i + 1), $entry.Value) -NoNewline
        $color = if ($status -eq 'Set') { 'Green' } else { 'Yellow' }
        Write-Host "[$status]" -ForegroundColor $color
    }
    Write-Host '  11. Test Connections' -ForegroundColor Cyan
    Write-Host '   0. Exit' -ForegroundColor Cyan

    $choice = (Read-Host -Prompt 'Choose a number, separate multiple numbers with commas, or press Enter for all').Trim()
    if ($choice -in @('0', 'Q')) { break }
    if ($choice -eq '11') {
        Show-ConnectionResults
        continue
    }

    $selected = if ($choice -eq '') {
        $options
    }
    else {
        $indices = [Collections.Generic.List[int]]::new()
        foreach ($number in ($choice -split ',')) {
            $index = 0
            if (![int]::TryParse($number.Trim(), [ref]$index) -or $index -lt 1 -or $index -gt $options.Count) {
                Write-Warning "Choose credential numbers from 1 to $($options.Count), enter 11 to test, or enter 0 to exit."
                continue MainMenu
            }
            $indices.Add($index - 1)
        }
        $indices | Select-Object -Unique | ForEach-Object { $options[$_] }
    }

    Write-Host 'Press Enter on a blank credential prompt to keep its current value.'
    foreach ($entry in $selected) {
        $value = Read-Host -Prompt $entry.Value -AsSecureString
        try {
            if ($value.Length -eq 0) { continue }
            $target = $targetPrefix + $entry.Key
            [BrickLinkLegacyCredentialStore]::Save($target, $value)
            Write-Host "Saved $($entry.Value)." -ForegroundColor Green
        }
        catch {
            Write-Warning "Could not save $($entry.Value). $($_.Exception.GetBaseException().Message)"
        }
        finally {
            $value.Dispose()
        }
    }
}
