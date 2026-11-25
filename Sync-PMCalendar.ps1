#Requires -Version 5.1
<#
.SYNOPSIS
    DF Kalender-Export & lokaler Outlook-Sync (Single-Config, DPAPI, dynamische URL, Workspace-ID & VPN-Check)
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# ============================================================================
# 0. Initialisierung & HtmlAgilityPack laden
# ============================================================================
$BaseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$HapDllPath = Join-Path $BaseDir "HtmlAgilityPack.dll"

if (-not (Test-Path $HapDllPath)) {
    Write-Host "FEHLER: 'HtmlAgilityPack.dll' wurde nicht im Skript-Ordner gefunden!" -ForegroundColor Red
    Write-Host "Bitte kopiere die HtmlAgilityPack.dll aus deinem C#-Projekt in: $BaseDir" -ForegroundColor Yellow
    Read-Host "`nDrücke Enter zum Beenden..."
    exit 1
}

Add-Type -Path $HapDllPath
Add-Type -AssemblyName System.Net.Http
Add-Type -AssemblyName System.Web

# ============================================================================
# 1. Hilfsfunktionen (VPN-Check, DPAPI-Verschlüsselung, Konfiguration & Zeit)
# ============================================================================
function Normalize-DFBaseUrl {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $null }

    $clean = $Url.Trim()
    if ($clean -notmatch '^https?://') {
        $clean = "https://$clean"
    }
    if (-not $clean.EndsWith("/")) {
        $clean = "$clean/"
    }

    [Uri]$parsedUri = $null
    if ([Uri]::TryCreate($clean, [UriKind]::Absolute, [ref]$parsedUri)) {
        return $parsedUri.AbsoluteUri
    }
    return $null
}

function Test-DFVpnIp {
    param([string]$BaseUrl)

    $normalizedUrl = Normalize-DFBaseUrl -Url $BaseUrl
    if ([string]::IsNullOrWhiteSpace($normalizedUrl)) {
        Write-Host "FEHLER: Ungültige URL '$BaseUrl'." -ForegroundColor Red
        return $false
    }

    $uri = [Uri]::new($normalizedUrl)
    $targetHost = $uri.Host

    Write-Host "`nPrüfe VPN-Verbindung (erwarte 10.x.x.x für '$targetHost')..."

    try {
        $addresses = [System.Net.Dns]::GetHostAddresses($targetHost)
        $ipv4List = @($addresses | Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork })

        if ($ipv4List.Count -eq 0) {
            Write-Host "WAN-Check fehlgeschlagen: Für '$targetHost' wurde keine IPv4-Adresse aufgelöst." -ForegroundColor Red
            return $false
        }

        $vpnIps = @($ipv4List | Where-Object { $_.GetAddressBytes()[0] -eq 10 })
        $allIpsText = ($ipv4List | ForEach-Object { $_.ToString() }) -join ", "

        if ($vpnIps.Count -gt 0) {
            $vpnIpsText = ($vpnIps | ForEach-Object { $_.ToString() }) -join ", "
            Write-Host "WAN-Check OK: '$targetHost' löst auf interne IP $vpnIpsText auf." -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "WAN-Check fehlgeschlagen: '$targetHost' löst auf '$allIpsText' auf (keine 10.x.x.x-Adresse)." -ForegroundColor Red
            Write-Host "Bitte zuerst mit dem Firmen-VPN bzw. Firmen-Netzwerk verbinden und das Skript erneut starten." -ForegroundColor Yellow
            return $false
        }
    }
    catch {
        Write-Host "WAN-Check fehlgeschlagen: DNS-Auflösung für '$targetHost' nicht möglich ($($_.Exception.Message))." -ForegroundColor Red
        Write-Host "Bitte zuerst mit dem Firmen-VPN bzw. Firmen-Netzwerk verbinden und das Skript erneut starten." -ForegroundColor Yellow
        return $false
    }
}

function Protect-PasswordDpapi {
    param([string]$PlainText)
    if ([string]::IsNullOrEmpty($PlainText)) { return $null }
    $secure = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
    return (ConvertFrom-SecureString -SecureString $secure)
}

function Unprotect-PasswordDpapi {
    param([string]$EncryptedText)
    if ([string]::IsNullOrWhiteSpace($EncryptedText)) { return $null }
    try {
        $secure = ConvertTo-SecureString -String $EncryptedText
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try {
            return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        }
        finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
    catch {
        Write-Host "Warnung: Gespeichertes Passwort konnte nicht entschlüsselt werden (anderer Benutzer/PC?)." -ForegroundColor Yellow
        return $null
    }
}

function Read-PasswordMasked {
    param([string]$Prompt = "Passwort: ")
    Write-Host -NoNewline $Prompt
    $sb = New-Object System.Text.StringBuilder
    while ($true) {
        $keyInfo = [Console]::ReadKey($true)
        if ($keyInfo.Key -eq [ConsoleKey]::Enter) {
            Write-Host ""
            break
        }
        elseif ($keyInfo.Key -eq [ConsoleKey]::Backspace) {
            if ($sb.Length -gt 0) {
                $sb.Remove($sb.Length - 1, 1) | Out-Null
                Write-Host -NoNewline "`b `b"
            }
        }
        elseif (-not [char]::IsControl($keyInfo.KeyChar)) {
            $sb.Append($keyInfo.KeyChar) | Out-Null
            Write-Host -NoNewline "*"
        }
    }
    return $sb.ToString()
}

function Save-AppConfig {
    param([string]$FilePath, [hashtable]$Config)

    $lines = [System.Collections.Generic.List[string]]::new()

    if (-not [string]::IsNullOrWhiteSpace($Config.BaseUrl)) {
        $lines.Add("BaseUrl=$($Config.BaseUrl)")
    }
    if (-not [string]::IsNullOrWhiteSpace($Config.WorkspaceId)) {
        $lines.Add("WorkspaceId=$($Config.WorkspaceId)")
    }

    $lines.Add("Username=$($Config.Username)")

    if (-not [string]::IsNullOrEmpty($Config.Password)) {
        $encryptedPw = Protect-PasswordDpapi -PlainText $Config.Password
        $lines.Add("EncryptedPassword=$encryptedPw")
    }

    if (-not [string]::IsNullOrWhiteSpace($Config.StartMonth)) {
        $lines.Add("StartMonth=$($Config.StartMonth)")
    }
    if (-not [string]::IsNullOrWhiteSpace($Config.EndMonth)) {
        $lines.Add("EndMonth=$($Config.EndMonth)")
    }
    if (-not [string]::IsNullOrWhiteSpace($Config.UserFilter)) {
        $lines.Add("UserFilter=$($Config.UserFilter)")
    }
    if (-not [string]::IsNullOrWhiteSpace($Config.OwnUserId)) {
        $lines.Add("OwnUserId=$($Config.OwnUserId)")
    }
    if (-not [string]::IsNullOrWhiteSpace($Config.OutlookCalendar)) {
        $lines.Add("OutlookCalendar=$($Config.OutlookCalendar)")
    }
    if ($null -ne $Config.CachedUsers -and $Config.CachedUsers.Count -gt 0) {
        $compactUsersJson = ConvertTo-Json -InputObject @($Config.CachedUsers) -Compress -Depth 4
        $lines.Add("CachedUsers=$compactUsersJson")
    }

    [System.IO.File]::WriteAllLines($FilePath, $lines, [System.Text.UTF8Encoding]::new($true))
}

function Format-MonthOffset {
    param([int]$Offset)
    if ($Offset -ge 0) { return "+$Offset" }
    return "$Offset"
}

function TryParse-MonthString {
    param([string]$InputStr, [ref]$ResultDate, [ref]$IsRelativeOffset)

    if ([string]::IsNullOrWhiteSpace($InputStr)) { return $false }
    $clean = $InputStr.Trim()

    $today = [datetime]::Today
    $currentMonthStart = [datetime]::new($today.Year, $today.Month, 1)

    if ($clean -match '^[+-]?\d{1,2}$') {
        $offset = [int]$clean
        $ResultDate.Value = $currentMonthStart.AddMonths($offset)
        if ($null -ne $IsRelativeOffset) { $IsRelativeOffset.Value = $true }
        return $true
    }

    [datetime]$parsed = [datetime]::MinValue
    foreach ($fmt in @("yyyy-MM", "MM.yyyy", "yyyy-MM-dd", "dd.MM.yyyy")) {
        if ([datetime]::TryParseExact($clean, $fmt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            $ResultDate.Value = [datetime]::new($parsed.Year, $parsed.Month, 1)
            if ($null -ne $IsRelativeOffset) { $IsRelativeOffset.Value = $false }
            return $true
        }
    }
    return $false
}

function Read-BaseUrlFromConsole {
    param([string]$DefaultUrl = "https://server.firma.de/XX/")

    while ($true) {
        $rawUrl = (Read-Host "URL [Enter = $DefaultUrl]").Trim()
        if ([string]::IsNullOrEmpty($rawUrl)) {
            return $DefaultUrl
        }

        $normalized = Normalize-DFBaseUrl -Url $rawUrl
        if (-not [string]::IsNullOrWhiteSpace($normalized)) {
            return $normalized
        }

        Write-Host "Ungültige URL. Bitte eine gültige Webadresse eingeben (z. B. https://server.firma.de/XX)." -ForegroundColor Yellow
    }
}

function Read-WorkspaceIdFromConsole {
    param([string]$DefaultWorkspaceId = "4DD936F6-33D2-4D56-AB08-B0D23FEE4B9D")

    $rawId = (Read-Host "Workspace-ID [Enter = $DefaultWorkspaceId]").Trim()
    if ([string]::IsNullOrEmpty($rawId)) {
        return $DefaultWorkspaceId
    }
    return $rawId
}

function Get-OrPromptAppConfig {
    param([string]$FilePath)

    $defaultBaseUrl     = "https://server.firma.de/XX/"
    $defaultWorkspaceId = "4DD936F6-33D2-4D56-AB08-B0D23FEE4B9D"

    $config = @{
        BaseUrl         = $null
        WorkspaceId     = $null
        Username        = $null
        Password        = $null
        StartMonth      = $null
        EndMonth        = $null
        UserFilter      = $null
        OwnUserId       = $null
        OutlookCalendar = $null
        CachedUsers     = @()
        PersistInFile   = $false
    }

    $hadPlaintextPassword = $false

    if (Test-Path $FilePath) {
        $config.PersistInFile = $true
        foreach ($rawLine in [System.IO.File]::ReadAllLines($FilePath, [System.Text.Encoding]::UTF8)) {
            $line = $rawLine.Trim()
            if ([string]::IsNullOrEmpty($line) -or $line.StartsWith("#")) { continue }

            $sepIdx = $line.IndexOf('=')
            if ($sepIdx -le 0) { continue }

            $key = $line.Substring(0, $sepIdx).Trim()
            $val = $line.Substring($sepIdx + 1).Trim()

            switch -Regex ($key) {
                '^BaseUrl$'           { $config.BaseUrl = Normalize-DFBaseUrl -Url $val }
                '^WorkspaceId$'       { $config.WorkspaceId = $val }
                '^Username$'          { $config.Username = $val }
                '^EncryptedPassword$' { $config.Password = Unprotect-PasswordDpapi -EncryptedText $val }
                '^Password$'          {
                    if ([string]::IsNullOrEmpty($config.Password) -and -not [string]::IsNullOrWhiteSpace($val)) {
                        $config.Password = $val
                        $hadPlaintextPassword = $true
                    }
                }
                '^StartMonth$'        { $config.StartMonth = $val }
                '^EndMonth$'          { $config.EndMonth = $val }
                '^UserFilter$'        { $config.UserFilter = $val }
                '^OwnUserId$'         { $config.OwnUserId = $val }
                '^OutlookCalendar$'   { $config.OutlookCalendar = $val }
                '^CachedUsers$'       {
                    if (-not [string]::IsNullOrWhiteSpace($val)) {
                        try {
                            $parsedUsers = @($val | ConvertFrom-Json)
                            $list = [System.Collections.Generic.List[object]]::new()
                            foreach ($u in $parsedUsers) {
                                if (-not [string]::IsNullOrWhiteSpace($u.userId)) {
                                    $list.Add([ordered]@{
                                        userId      = [string]$u.userId
                                        displayName = [string]$u.displayName
                                        isSelected  = [bool]$u.isSelected
                                    })
                                }
                            }
                            $config.CachedUsers = $list.ToArray()
                        }
                        catch {
                            Write-Host "Warnung: 'CachedUsers' in 'credentials.txt' konnte nicht gelesen werden und wird neu geladen." -ForegroundColor Yellow
                        }
                    }
                }
            }
        }
    }

    $hasValidCredentials = (
        -not [string]::IsNullOrWhiteSpace($config.Username) -and
        -not [string]::IsNullOrWhiteSpace($config.Password) -and
        $config.Username -ne "DEIN_BENUTZERNAME"
    )

    if ($hasValidCredentials) {
        $needsSave = $hadPlaintextPassword

        if ([string]::IsNullOrWhiteSpace($config.BaseUrl)) {
            Write-Host "In 'credentials.txt' ist noch keine URL hinterlegt."
            $config.BaseUrl = Read-BaseUrlFromConsole -DefaultUrl $defaultBaseUrl
            $needsSave = $true
        }

        if ([string]::IsNullOrWhiteSpace($config.WorkspaceId)) {
            Write-Host "In 'credentials.txt' ist noch keine Workspace-ID hinterlegt."
            $config.WorkspaceId = Read-WorkspaceIdFromConsole -DefaultWorkspaceId $defaultWorkspaceId
            $needsSave = $true
        }

        if ($needsSave -and $config.PersistInFile) {
            Save-AppConfig -FilePath $FilePath -Config $config
            Write-Host "Fehlende Konfigurationswerte aktualisiert und in 'credentials.txt' gespeichert." -ForegroundColor Green
        }

        Write-Host "Konfiguration aus 'credentials.txt' geladen:"
        Write-Host "  URL:                        $($config.BaseUrl)"
        Write-Host "  Workspace-ID:               $($config.WorkspaceId)"
        Write-Host "  Benutzername:               $($config.Username) (Passwort: [DPAPI-verschlüsselt])"
        if ($config.StartMonth -or $config.EndMonth) {
            [datetime]$sDt = [datetime]::MinValue
            [datetime]$eDt = [datetime]::MinValue
            $sRel = $false
            $eRel = $false
            $okS = TryParse-MonthString -InputStr $config.StartMonth -ResultDate ([ref]$sDt) -IsRelativeOffset ([ref]$sRel)
            $okE = TryParse-MonthString -InputStr $config.EndMonth   -ResultDate ([ref]$eDt) -IsRelativeOffset ([ref]$eRel)

            if ($okS -and $okE -and ($sRel -or $eRel)) {
                Write-Host "  Gespeicherter Zeitraum:     $($config.StartMonth) bis $($config.EndMonth) Monate (gleitend -> $($sDt.ToString('yyyy-MM')) bis $($eDt.ToString('yyyy-MM')))"
            } else {
                Write-Host "  Gespeicherter Zeitraum:     $($config.StartMonth) bis $($config.EndMonth)"
            }
        }
        if ($config.UserFilter)              { Write-Host "  Gespeicherter Benutzerfilter: $($config.UserFilter)" }
        if ($config.OwnUserId)               { Write-Host "  Eigener DF-Kalender:          $($config.OwnUserId)" }
        if ($config.OutlookCalendar)         { Write-Host "  Ziel-Kalender in Outlook:     $($config.OutlookCalendar)" }
        if ($config.CachedUsers.Count -gt 0) { Write-Host "  Zwischengespeicherte User:    $($config.CachedUsers.Count) Einträge in 'credentials.txt'" }
        return $config
    }

    Write-Host "Ersteinrichtung: Bitte DF-Verbindungsdaten eingeben:`n"

    if ([string]::IsNullOrWhiteSpace($config.BaseUrl)) {
        $config.BaseUrl = Read-BaseUrlFromConsole -DefaultUrl $defaultBaseUrl
    } else {
        Write-Host "URL: $($config.BaseUrl)"
    }

    if ([string]::IsNullOrWhiteSpace($config.WorkspaceId)) {
        $config.WorkspaceId = Read-WorkspaceIdFromConsole -DefaultWorkspaceId $defaultWorkspaceId
    } else {
        Write-Host "Workspace-ID: $($config.WorkspaceId)"
    }

    if ([string]::IsNullOrWhiteSpace($config.Username)) {
        $config.Username = (Read-Host "Benutzername").Trim()
    } else {
        Write-Host "Benutzername: $($config.Username)"
    }

    $config.Password = Read-PasswordMasked -Prompt "Passwort: "

    if (-not [string]::IsNullOrWhiteSpace($config.Username) -and -not [string]::IsNullOrWhiteSpace($config.Password)) {
        $saveChoice = (Read-Host "URL, Workspace-ID, Zugangsdaten (Passwort verschlüsselt) & Einstellungen in 'credentials.txt' speichern? (j/n) [j]").Trim()
        $config.PersistInFile = [string]::IsNullOrEmpty($saveChoice) -or $saveChoice -match '^(j|y)$'

        if ($config.PersistInFile) {
            Save-AppConfig -FilePath $FilePath -Config $config
            Write-Host "Konfiguration verschlüsselt gespeichert in: $FilePath" -ForegroundColor Green
        }
    }

    return $config
}

function Read-MonthFromConsole {
    param([string]$Prompt, [datetime]$DefaultValue)

    while ($true) {
        Write-Host -NoNewline $Prompt
        $inputStr = (Read-Host).Trim()

        if ([string]::IsNullOrEmpty($inputStr)) {
            return $DefaultValue
        }

        [datetime]$parsed = [datetime]::MinValue
        $dummyRel = $false
        if (TryParse-MonthString -InputStr $inputStr -ResultDate ([ref]$parsed) -IsRelativeOffset ([ref]$dummyRel)) {
            return $parsed
        }

        Write-Host "Ungültiges Format. Bitte JJJJ-MM (z. B. 2026-11), MM.JJJJ oder relativen Offset (z. B. +0, +1) eingeben." -ForegroundColor Yellow
    }
}

function Get-OrPromptMonthsToFetch {
    param([hashtable]$Config, [string]$CredentialsPath)

    $today = [datetime]::Today
    $defaultStart = [datetime]::new($today.Year, $today.Month, 1)
    $defaultEnd = $defaultStart.AddMonths(1)

    [datetime]$startMonth = [datetime]::MinValue
    [datetime]$endMonth   = [datetime]::MinValue
    $isStartRel = $false
    $isEndRel   = $false

    $hasValidStart = TryParse-MonthString -InputStr $Config.StartMonth -ResultDate ([ref]$startMonth) -IsRelativeOffset ([ref]$isStartRel)
    $hasValidEnd   = TryParse-MonthString -InputStr $Config.EndMonth   -ResultDate ([ref]$endMonth)   -IsRelativeOffset ([ref]$isEndRel)

    if ($hasValidStart -and -not $hasValidEnd) {
        $endMonth = $startMonth
        $hasValidEnd = $true
    }

    if ($hasValidStart -and $hasValidEnd) {
        if ($endMonth -lt $startMonth) {
            Write-Host "Hinweis: Gespeicherter Endmonat liegt vor Startmonat - frage nur $($startMonth.ToString('yyyy-MM')) ab." -ForegroundColor Yellow
            $endMonth = $startMonth
        }
    }
    else {
        if (-not [string]::IsNullOrWhiteSpace($Config.StartMonth) -or -not [string]::IsNullOrWhiteSpace($Config.EndMonth)) {
            Write-Host "Warnung: Gespeicherter Zeitraum in 'credentials.txt' hat ein ungültiges Format. Bitte neu eingeben." -ForegroundColor Yellow
        }

        Write-Host "`nZeitraum für den Kalender-Abruf festlegen (Format: JJJJ-MM wie 2026-11 oder relativ wie +0 / +1):"

        $startMonth = Read-MonthFromConsole -Prompt "Startmonat [Enter = $($defaultStart.ToString('yyyy-MM'))]: " -DefaultValue $defaultStart
        $defaultEndForStart = if ($startMonth -gt $defaultEnd) { $startMonth } else { $defaultEnd }
        $endMonth = Read-MonthFromConsole -Prompt "Endmonat   [Enter = $($defaultEndForStart.ToString('yyyy-MM'))]: " -DefaultValue $defaultEndForStart

        if ($endMonth -lt $startMonth) {
            Write-Host "Hinweis: Endmonat liegt vor Startmonat - frage nur $($startMonth.ToString('yyyy-MM')) ab." -ForegroundColor Yellow
            $endMonth = $startMonth
        }

        if ($Config.PersistInFile -or (Test-Path $CredentialsPath)) {
            $startOffset = (($startMonth.Year - $defaultStart.Year) * 12) + ($startMonth.Month - $defaultStart.Month)
            $endOffset   = (($endMonth.Year   - $defaultStart.Year) * 12) + ($endMonth.Month   - $defaultStart.Month)
            $startOffStr = Format-MonthOffset $startOffset
            $endOffStr   = Format-MonthOffset $endOffset

            Write-Host "`nWie soll dieser Zeitraum in 'credentials.txt' gespeichert werden?"
            Write-Host "  [1] Gleitend / dynamisch ($startOffStr bis $endOffStr Monate - wandert jeden Monat automatisch mit)"
            Write-Host "  [2] Fester Zeitraum ($($startMonth.ToString('yyyy-MM')) bis $($endMonth.ToString('yyyy-MM')))"
            Write-Host "  [3] Nicht speichern (beim nächsten Start erneut fragen)"

            while ($true) {
                $saveChoice = (Read-Host "Auswahl [Enter = 1 (Gleitend)]").Trim()
                if ([string]::IsNullOrEmpty($saveChoice) -or $saveChoice -eq "1") {
                    $Config.StartMonth = $startOffStr
                    $Config.EndMonth   = $endOffStr
                    Save-AppConfig -FilePath $CredentialsPath -Config $Config
                    Write-Host "Gleitender Zeitraum ($startOffStr bis $endOffStr) in 'credentials.txt' gespeichert."
                    break
                }
                elseif ($saveChoice -eq "2") {
                    $Config.StartMonth = $startMonth.ToString("yyyy-MM")
                    $Config.EndMonth   = $endMonth.ToString("yyyy-MM")
                    Save-AppConfig -FilePath $CredentialsPath -Config $Config
                    Write-Host "Fester Zeitraum ($($Config.StartMonth) bis $($Config.EndMonth)) in 'credentials.txt' gespeichert."
                    break
                }
                elseif ($saveChoice -eq "3") {
                    break
                }
                Write-Host "Ungültige Eingabe. Bitte 1, 2 oder 3 wählen." -ForegroundColor Yellow
            }
        }
    }

    $months = [System.Collections.Generic.List[datetime]]::new()
    for ($curr = $startMonth; $curr -le $endMonth; $curr = $curr.AddMonths(1)) {
        $months.Add($curr)
    }

    $formattedList = ($months | ForEach-Object { $_.ToString("yyyy-MM") }) -join ", "
    Write-Host "`nAusgewählter Zeitraum: $formattedList"
    return $months
}

function Save-TextToFile {
    param([string]$FilePath, [string]$Content, [string]$Label)

    $dir = [System.IO.Path]::GetDirectoryName($FilePath)
    if (-not [string]::IsNullOrEmpty($dir) -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    [System.IO.File]::WriteAllText($FilePath, $Content, [System.Text.UTF8Encoding]::new($false))
    Write-Host "$Label erfolgreich gespeichert ($($Content.Length) Zeichen):"
    Write-Host "$FilePath`n"
}

# ============================================================================
# 2. DF HTTP-Client Funktionen
# ============================================================================
function New-DFSession {
    param([string]$BaseUrl = "https://server.firma.de/XX/")

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

    $normalizedUrl = Normalize-DFBaseUrl -Url $BaseUrl
    if ([string]::IsNullOrWhiteSpace($normalizedUrl)) {
        throw "Ungültige DF BaseUrl: '$BaseUrl'"
    }

    $baseUri = [Uri]::new($normalizedUrl)
    $cookieContainer = [System.Net.CookieContainer]::new()
    $handler = [System.Net.Http.HttpClientHandler]::new()
    $handler.CookieContainer = $cookieContainer
    $handler.UseCookies = $true
    $handler.AllowAutoRedirect = $true

    $client = [System.Net.Http.HttpClient]::new($handler)
    $client.BaseAddress = $baseUri
    $client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
    $client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest")

    return [PSCustomObject]@{
        Client          = $client
        CookieContainer = $cookieContainer
        BaseUri         = $baseUri
    }
}

function Connect-DF {
    param(
        $Session,
        [string]$Username,
        [string]$Password,
        [string]$WorkspaceId = "4DD936F6-33D2-4D56-AB08-B0D23FEE4B9D"
    )

    $Session.Client.GetAsync("").GetAwaiter().GetResult() | Out-Null

    $kvpList = [System.Collections.Generic.List[System.Collections.Generic.KeyValuePair[string, string]]]::new()
    $kvpList.Add([System.Collections.Generic.KeyValuePair[string, string]]::new("UserName", $Username))
    $kvpList.Add([System.Collections.Generic.KeyValuePair[string, string]]::new("Password", $Password))
    $kvpList.Add([System.Collections.Generic.KeyValuePair[string, string]]::new("LogOn", "true"))

    $formData = [System.Net.Http.FormUrlEncodedContent]::new($kvpList)
    $loginResp = $Session.Client.PostAsync("Environment/Scripting/WebRequest.php", $formData).GetAwaiter().GetResult()

    if (-not $loginResp.IsSuccessStatusCode) {
        Write-Host "Login-POST fehlgeschlagen: $([int]$loginResp.StatusCode) $($loginResp.ReasonPhrase)" -ForegroundColor Red
        return $false
    }

    $encodedWsId = [System.Uri]::EscapeDataString($WorkspaceId.Trim())
    $wsResp = $Session.Client.GetAsync("Environment/Workspace/Workspace.php?id=$encodedWsId").GetAwaiter().GetResult()
    if (-not $wsResp.IsSuccessStatusCode) {
        Write-Host "Workspace-Aufruf fehlgeschlagen: $([int]$wsResp.StatusCode) $($wsResp.ReasonPhrase)" -ForegroundColor Red
        return $false
    }

    $html = $wsResp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    return ($html.Contains("Workspace_Load(") -and $html.Contains("SessionTimeOut"))
}

function Get-DFPhpSessionId {
    param($Session)
    $cookies = $Session.CookieContainer.GetCookies($Session.BaseUri)
    $sessCookie = $cookies["PHPSESSID"]
    if ($sessCookie) { return $sessCookie.Value }
    return "Keine"
}

function Get-DFUsersDropdownHtml {
    param($Session)
    $resp = $Session.Client.GetAsync("Content/widgets/common/item/dropdown_selection_user.php").GetAwaiter().GetResult()
    if (-not $resp.IsSuccessStatusCode) {
        Write-Host "User-Dropdown-Abruf fehlgeschlagen: $([int]$resp.StatusCode) $($resp.ReasonPhrase)" -ForegroundColor Red
        return $null
    }
    return $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
}

function Get-DFCalendarMonthHtml {
    param($Session, [string]$DateStamp, [string]$UsersJsonArray, [int]$ScrollValue = 0)

    $kvpList = [System.Collections.Generic.List[System.Collections.Generic.KeyValuePair[string, string]]]::new()
    $kvpList.Add([System.Collections.Generic.KeyValuePair[string, string]]::new("DateStamp", $DateStamp))
    $kvpList.Add([System.Collections.Generic.KeyValuePair[string, string]]::new("Users", $UsersJsonArray))
    $kvpList.Add([System.Collections.Generic.KeyValuePair[string, string]]::new("ScrollValue", $ScrollValue.ToString()))

    $postData = [System.Net.Http.FormUrlEncodedContent]::new($kvpList)
    $resp = $Session.Client.PostAsync("Content/Scheduling/calendar_month.php?", $postData).GetAwaiter().GetResult()

    if (-not $resp.IsSuccessStatusCode) {
        Write-Host "Kalender-Abruf fehlgeschlagen: $([int]$resp.StatusCode) $($resp.ReasonPhrase)" -ForegroundColor Red
        return $null
    }
    return $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
}

function Get-DFModalHtml {
    param($Session, [string]$Category, [string]$DbId)

    $endpoint = switch ($Category.ToUpperInvariant()) {
        "MEETING"     { "Environment//Modal/Calendar_Meeting.php?" }
        "ABSENCE"     { "Environment//Modal/Calendar_Absence.php?" }
        "SPECIALTIME" { "Environment//Modal/Calendar_Specialtime.php?" }
        default       { return $null }
    }

    $kvpList = [System.Collections.Generic.List[System.Collections.Generic.KeyValuePair[string, string]]]::new()
    $kvpList.Add([System.Collections.Generic.KeyValuePair[string, string]]::new("Mode", "4"))
    $kvpList.Add([System.Collections.Generic.KeyValuePair[string, string]]::new("DB_ID", $DbId))

    $postData = [System.Net.Http.FormUrlEncodedContent]::new($kvpList)
    $resp = $Session.Client.PostAsync($endpoint, $postData).GetAwaiter().GetResult()

    if (-not $resp.IsSuccessStatusCode) { return $null }
    return $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
}

function Get-DFMeetingHistoryHtml {
    param($Session, [string]$DbId)

    $kvpList = [System.Collections.Generic.List[System.Collections.Generic.KeyValuePair[string, string]]]::new()
    $kvpList.Add([System.Collections.Generic.KeyValuePair[string, string]]::new("DB_ID", $DbId))

    $postData = [System.Net.Http.FormUrlEncodedContent]::new($kvpList)
    $resp = $Session.Client.PostAsync("Environment//Modal/Calendar_History_Meeting.php?", $postData).GetAwaiter().GetResult()

    if (-not $resp.IsSuccessStatusCode) { return $null }
    return $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
}

# ============================================================================
# 3. HTML-Parser Hilfsfunktionen (HtmlAgilityPack)
# ============================================================================
function Get-HtmlAttribute {
    param([HtmlAgilityPack.HtmlDocument]$Doc, [string]$Id, [string]$Attr)
    $node = $Doc.DocumentNode.SelectSingleNode("//*[@id='$Id']")
    if ($null -eq $node) { return $null }
    return [HtmlAgilityPack.HtmlEntity]::DeEntitize($node.GetAttributeValue($Attr, "")).Trim()
}

function Get-HtmlInnerTextById {
    param([HtmlAgilityPack.HtmlDocument]$Doc, [string]$Id)
    $node = $Doc.DocumentNode.SelectSingleNode("//*[@id='$Id']")
    if ($null -eq $node) { return $null }
    return [HtmlAgilityPack.HtmlEntity]::DeEntitize($node.InnerText).Trim()
}

function Get-HtmlSelectedOption {
    param([HtmlAgilityPack.HtmlDocument]$Doc, [string]$SelectId)
    $opt = $Doc.DocumentNode.SelectSingleNode("//select[@id='$SelectId']/option[@selected]")
    if ($null -eq $opt) {
        return @{ Value = $null; Text = $null }
    }
    return @{
        Value = $opt.GetAttributeValue("value", "").Trim()
        Text  = [HtmlAgilityPack.HtmlEntity]::DeEntitize($opt.InnerText).Trim()
    }
}

function Get-HtmlItemTextLines {
    param([HtmlAgilityPack.HtmlDocument]$Doc, [string]$FieldId)
    $idNode = $Doc.DocumentNode.SelectSingleNode("//*[@id='$FieldId']")
    if ($null -eq $idNode) { return @() }

    $itemContainer = $null
    foreach ($anc in $idNode.Ancestors()) {
        if ($anc.HasClass("item")) {
            $itemContainer = $anc
            break
        }
    }
    if ($null -eq $itemContainer) { return @() }

    $textNodes = $itemContainer.SelectNodes(".//*[contains(@class, 'textline')]")
    if ($null -eq $textNodes) { return @() }

    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($n in $textNodes) {
        $txt = [HtmlAgilityPack.HtmlEntity]::DeEntitize($n.InnerText).Trim()
        if (-not [string]::IsNullOrWhiteSpace($txt)) {
            $lines.Add($txt)
        }
    }
    return $lines.ToArray()
}

function Normalize-TimeString {
    param([string]$TimeVal)
    if ([string]::IsNullOrWhiteSpace($TimeVal)) { return $null }
    [timespan]$ts = [timespan]::Zero
    if ([timespan]::TryParse($TimeVal, [ref]$ts)) {
        return $ts.ToString("hh\:mm")
    }
    return $TimeVal
}

function Parse-HtmlTimePeriod {
    param(
        [HtmlAgilityPack.HtmlDocument]$Doc,
        [string]$DateFromId,
        [string]$TimeFromId,
        [string]$TimeUntilId,
        $FallbackEntry = $null
    )

    $dateFromNode  = $Doc.DocumentNode.SelectSingleNode("//*[@id='$DateFromId']")
    $timeFromNode  = $Doc.DocumentNode.SelectSingleNode("//*[@id='$TimeFromId']")
    $timeUntilNode = $Doc.DocumentNode.SelectSingleNode("//*[@id='$TimeUntilId']")

    $rawDateFrom  = if ($dateFromNode)  { [HtmlAgilityPack.HtmlEntity]::DeEntitize($dateFromNode.GetAttributeValue("value", "")).Trim() } else { $null }
    $rawTimeFrom  = if ($timeFromNode)  { [HtmlAgilityPack.HtmlEntity]::DeEntitize($timeFromNode.GetAttributeValue("value", "")).Trim() } else { $null }
    $rawTimeUntil = if ($timeUntilNode) { [HtmlAgilityPack.HtmlEntity]::DeEntitize($timeUntilNode.GetAttributeValue("value", "")).Trim() } else { $null }

    $minTime = if ($timeFromNode) { $timeFromNode.GetAttributeValue("min", $null) } else { $null }
    $maxTime = if ($timeFromNode) { $timeFromNode.GetAttributeValue("max", $null) } else { $null }

    [int]$stepOut = 0
    $stepSeconds = if ($timeFromNode -and [int]::TryParse($timeFromNode.GetAttributeValue("step", ""), [ref]$stepOut)) { $stepOut } else { $null }

    $normTimeFrom  = Normalize-TimeString $rawTimeFrom
    $normTimeUntil = Normalize-TimeString $rawTimeUntil

    $resolvedDateFrom = if (-not [string]::IsNullOrWhiteSpace($rawDateFrom)) { $rawDateFrom } elseif ($FallbackEntry) { $FallbackEntry.date } else { $null }

    $resolvedTimeFrom = if (-not [string]::IsNullOrWhiteSpace($normTimeFrom)) {
        $normTimeFrom
    } elseif ($FallbackEntry -and $FallbackEntry.startDateTime -and $FallbackEntry.startDateTime.Contains(" ")) {
        $FallbackEntry.startDateTime.Split(' ')[1]
    } else {
        $null
    }

    $resolvedTimeUntil = if (-not [string]::IsNullOrWhiteSpace($normTimeUntil)) {
        $normTimeUntil
    } elseif ($FallbackEntry -and $FallbackEntry.endDateTime -and $FallbackEntry.endDateTime.Contains(" ")) {
        $FallbackEntry.endDateTime.Split(' ')[1]
    } else {
        $null
    }

    $startDateTime = if (-not [string]::IsNullOrWhiteSpace($resolvedDateFrom) -and -not [string]::IsNullOrWhiteSpace($resolvedTimeFrom)) {
        "$resolvedDateFrom $resolvedTimeFrom"
    } elseif ($FallbackEntry) {
        $FallbackEntry.startDateTime
    } else {
        $null
    }

    $endDateTime = if (-not [string]::IsNullOrWhiteSpace($resolvedDateFrom) -and -not [string]::IsNullOrWhiteSpace($resolvedTimeUntil)) {
        "$resolvedDateFrom $resolvedTimeUntil"
    } elseif ($FallbackEntry) {
        $FallbackEntry.endDateTime
    } else {
        $null
    }

    $durationHours = if ($FallbackEntry) { $FallbackEntry.durationHours } else { $null }
    [timespan]$tFrom = [timespan]::Zero
    [timespan]$tUntil = [timespan]::Zero
    if ([timespan]::TryParse($resolvedTimeFrom, [ref]$tFrom) -and
        [timespan]::TryParse($resolvedTimeUntil, [ref]$tUntil) -and
        $tUntil -ge $tFrom) {
        $durationHours = [Math]::Round(($tUntil - $tFrom).TotalHours, 2)
    }

    return [ordered]@{
        label             = "Zeitraum (von - bis)"
        htmlDateFrom      = $rawDateFrom
        htmlTimeFrom      = $rawTimeFrom
        htmlTimeUntil     = $rawTimeUntil
        resolvedDateFrom  = $resolvedDateFrom
        resolvedTimeFrom  = $resolvedTimeFrom
        resolvedTimeUntil = $resolvedTimeUntil
        startDateTime     = $startDateTime
        endDateTime       = $endDateTime
        durationHours     = $durationHours
        minTime           = $minTime
        maxTime           = $maxTime
        stepSeconds       = $stepSeconds
    }
}

# ============================================================================
# 4. Spezifische Parser (Users, Calendar, Modals, History)
# ============================================================================
function Parse-UserDropdown {
    param([string]$Html)

    $doc = [HtmlAgilityPack.HtmlDocument]::new()
    $doc.LoadHtml($Html)

    $result = [System.Collections.Generic.List[object]]::new()
    $itemNodes = $doc.DocumentNode.SelectNodes("//*[contains(@class, 'assignment_dropdown-item') and @id]")
    if ($null -eq $itemNodes) { return $result }

    foreach ($itemNode in $itemNodes) {
        $userId = $itemNode.GetAttributeValue("id", "").Trim()
        if ([string]::IsNullOrWhiteSpace($userId)) { continue }

        $nameNode = $itemNode.SelectSingleNode(".//*[@data-id]")
        $displayName = [HtmlAgilityPack.HtmlEntity]::DeEntitize($(if ($nameNode) { $nameNode.InnerText } else { "" })).Trim()

        $checkboxNode = $itemNode.SelectSingleNode(".//input[@type='checkbox']")
        $hasCheckedAttr = ($null -ne $checkboxNode) -and ($null -ne $checkboxNode.Attributes["checked"])

        $inSelectedContainer = $false
        foreach ($anc in $itemNode.Ancestors()) {
            if ($anc.HasClass("selected")) {
                $inSelectedContainer = $true
                break
            }
        }

        $result.Add([ordered]@{
            userId      = $userId
            displayName = $displayName
            isSelected  = ($hasCheckedAttr -or $inSelectedContainer)
        })
    }

    return $result
}

function Parse-MeetingModal {
    param([string]$ModalHtml, $FallbackEntry = $null)

    $doc = [HtmlAgilityPack.HtmlDocument]::new()
    $doc.LoadHtml($ModalHtml)

    $tp  = Parse-HtmlTimePeriod -Doc $doc -DateFromId "Meeting_DateFrom" -TimeFromId "Meeting_TimeFrom" -TimeUntilId "Meeting_TimeUntil" -FallbackEntry $FallbackEntry
    $op  = Get-HtmlSelectedOption -Doc $doc -SelectId "ID_Operation"
    $use = Get-HtmlSelectedOption -Doc $doc -SelectId "Use"

    return [ordered]@{
        mode             = Get-HtmlAttribute -Doc $doc -Id "Mode_Meeting" -Attr "value"
        status           = Get-HtmlAttribute -Doc $doc -Id "Status_Meeting" -Attr "value"
        customerId       = Get-HtmlAttribute -Doc $doc -Id "ID_Customer" -Attr "value"
        customerLines    = @(Get-HtmlItemTextLines -Doc $doc -FieldId "ID_Customer")
        contactId        = Get-HtmlAttribute -Doc $doc -Id "ID_Contact" -Attr "value"
        contactLines     = @(Get-HtmlItemTextLines -Doc $doc -FieldId "ID_Contact")
        meetingUserId    = Get-HtmlAttribute -Doc $doc -Id "ID_User_Meeting" -Attr "value"
        meetingUserLines = @(Get-HtmlItemTextLines -Doc $doc -FieldId "ID_User_Meeting")
        orderId          = Get-HtmlAttribute -Doc $doc -Id "ID_Order" -Attr "value"
        orderLines       = @(Get-HtmlItemTextLines -Doc $doc -FieldId "ID_Order")
        ticket           = Get-HtmlInnerTextById -Doc $doc -Id "Meeting_Ticket"
        operationId      = $op.Value
        operationName    = $op.Text
        useId            = $use.Value
        useName          = $use.Text
        meetingDateFrom  = $tp.htmlDateFrom
        meetingTimeFrom  = $tp.htmlTimeFrom
        meetingTimeUntil = $tp.htmlTimeUntil
        timePeriod       = $tp
        description      = Get-HtmlInnerTextById -Doc $doc -Id "Meeting_Description"
        history          = @()
    }
}

function Parse-MeetingHistory {
    param([string]$HistoryHtml)

    $result = [System.Collections.Generic.List[object]]::new()
    if ([string]::IsNullOrWhiteSpace($HistoryHtml)) { return @() }

    $doc = [HtmlAgilityPack.HtmlDocument]::new()
    $doc.LoadHtml($HistoryHtml)

    $headerItems = $doc.DocumentNode.SelectNodes("//div[contains(@class, 'widget_content')]/div[contains(@class, 'item') and not(contains(@class, 'expcontent'))]")
    if ($null -eq $headerItems) { return @() }

    foreach ($headerItem in $headerItems) {
        $cells = $headerItem.SelectNodes(".//*[contains(@class, 'wscell')]")
        if ($null -eq $cells -or $cells.Count -lt 3) { continue }

        $author       = [HtmlAgilityPack.HtmlEntity]::DeEntitize($cells[1].InnerText).Trim()
        $timestampRaw = [HtmlAgilityPack.HtmlEntity]::DeEntitize($cells[2].InnerText).Trim()

        $timestampIso = $null
        $tsMatch = [regex]::Match($timestampRaw, '(\d{2}\.\d{2}\.\d{4})\s*\|\s*(\d{2}:\d{2})')
        if ($tsMatch.Success) {
            [datetime]$dt = [datetime]::MinValue
            if ([datetime]::TryParseExact("$($tsMatch.Groups[1].Value) $($tsMatch.Groups[2].Value)", "dd.MM.yyyy HH:mm", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
                $timestampIso = $dt.ToString("yyyy-MM-dd HH:mm")
            }
        }

        $entry = [ordered]@{
            author       = $author
            timestampRaw = $timestampRaw
            timestampIso = $timestampIso
            status       = $null
            description  = $null
            fields       = [ordered]@{}
        }

        $contentItem = $headerItem.NextSibling
        while ($null -ne $contentItem -and $contentItem.NodeType -ne [HtmlAgilityPack.HtmlNodeType]::Element) {
            $contentItem = $contentItem.NextSibling
        }

        if ($null -ne $contentItem -and $contentItem.HasClass("expcontent")) {
            $labelNodes = $contentItem.SelectNodes("./div[contains(@class, 'lb')]")
            if ($null -ne $labelNodes) {
                foreach ($lbNode in $labelNodes) {
                    $key = [HtmlAgilityPack.HtmlEntity]::DeEntitize($lbNode.InnerText).Trim()

                    $valNode = $lbNode.NextSibling
                    while ($null -ne $valNode -and $valNode.NodeType -ne [HtmlAgilityPack.HtmlNodeType]::Element) {
                        $valNode = $valNode.NextSibling
                    }

                    if ($null -ne $valNode -and $valNode.HasClass("textline")) {
                        $htmlWithNewlines = [regex]::Replace($valNode.InnerHtml, '<br\s*/?>', "`n", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                        $tempDoc = [HtmlAgilityPack.HtmlDocument]::new()
                        $tempDoc.LoadHtml($htmlWithNewlines)
                        $decoded = [HtmlAgilityPack.HtmlEntity]::DeEntitize($tempDoc.DocumentNode.InnerText)
                        $lines = $decoded.Replace("`r`n", "`n").Split("`n") | ForEach-Object { $_.Trim() }
                        $value = ($lines -join "`n").Trim()

                        if (-not [string]::IsNullOrEmpty($key)) {
                            $entry.fields[$key] = $value
                            if ($key.Equals("Status", [StringComparison]::OrdinalIgnoreCase)) {
                                $entry.status = $value
                            }
                            elseif ($key.Equals("Beschreibung", [StringComparison]::OrdinalIgnoreCase)) {
                                $entry.description = $value
                            }
                        }
                    }
                }
            }
        }

        $result.Add($entry)
    }

    return $result.ToArray()
}

function Parse-AbsenceModal {
    param([string]$ModalHtml)

    $doc = [HtmlAgilityPack.HtmlDocument]::new()
    $doc.LoadHtml($ModalHtml)

    $person = Get-HtmlSelectedOption -Doc $doc -SelectId "ID_User"
    $type   = Get-HtmlSelectedOption -Doc $doc -SelectId "Absence_Type"
    $reason = Get-HtmlSelectedOption -Doc $doc -SelectId "ID_Absence"

    return [ordered]@{
        modeAbsence       = Get-HtmlAttribute -Doc $doc -Id "Mode_Absence" -Attr "value"
        personUserId      = $person.Value
        personName        = $person.Text
        dateStart         = Get-HtmlAttribute -Doc $doc -Id "Absence_DateStart" -Attr "value"
        dateStop          = Get-HtmlAttribute -Doc $doc -Id "Absence_DateStop" -Attr "value"
        absenceTypeId     = $type.Value
        absenceTypeName   = $type.Text
        absenceReasonId   = $reason.Value
        absenceReasonName = $reason.Text
        description       = Get-HtmlInnerTextById -Doc $doc -Id "Absence_Description"
    }
}

function Parse-SpecialtimeModal {
    param([string]$ModalHtml, $FallbackEntry = $null)

    $doc = [HtmlAgilityPack.HtmlDocument]::new()
    $doc.LoadHtml($ModalHtml)

    $person = Get-HtmlSelectedOption -Doc $doc -SelectId "Specialtime_ID_User"
    $type   = Get-HtmlSelectedOption -Doc $doc -SelectId "ID_Specialtime"
    $tp     = Parse-HtmlTimePeriod -Doc $doc -DateFromId "Specialtime_DateFrom" -TimeFromId "Specialtime_TimeFrom" -TimeUntilId "Specialtime_TimeUntil" -FallbackEntry $FallbackEntry

    return [ordered]@{
        modeSpecialtime      = Get-HtmlAttribute -Doc $doc -Id "Mode_Specialtime" -Attr "value"
        personUserId         = $person.Value
        personName           = $person.Text
        specialtimeDateFrom  = $tp.htmlDateFrom
        specialtimeTimeFrom  = $tp.htmlTimeFrom
        specialtimeTimeUntil = $tp.htmlTimeUntil
        timePeriod           = $tp
        specialtimeTypeId    = $type.Value
        specialtimeTypeName  = $type.Text
        description          = Get-HtmlInnerTextById -Doc $doc -Id "Specialtime_Description"
    }
}

function Extract-CssProperty {
    param([string]$Style, [string]$PropertyName)
    if ([string]::IsNullOrWhiteSpace($Style)) { return $null }
    $pattern = "(?:^|;)\s*$([regex]::Escape($PropertyName))\s*:\s*([^;]+)"
    $match = [regex]::Match($Style, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($match.Success) { return $match.Groups[1].Value.Trim() }
    return $null
}

function Parse-PxValue {
    param([string]$CssVal)
    if ([string]::IsNullOrWhiteSpace($CssVal)) { return $null }
    $match = [regex]::Match($CssVal, '([-+]?[\d.,]+)\s*px', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($match.Success) {
        [double]$val = 0
        if ([double]::TryParse($match.Groups[1].Value.Replace(',', '.'), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$val)) {
            return $val
        }
    }
    return $null
}

function Find-CssPxInHierarchy {
    param($StartNode, $StopAtNode, [string]$PropertyName)
    $current = $StartNode
    while ($null -ne $current -and $current -ne $StopAtNode) {
        $style = $current.GetAttributeValue("style", "")
        $val = Parse-PxValue (Extract-CssProperty -Style $style -PropertyName $PropertyName)
        if ($null -ne $val) { return $val }
        $current = $current.ParentNode
    }
    return $null
}

function Parse-CalendarHtml {
    param([string]$HtmlContent, [string]$PeriodStart)

    $doc = [HtmlAgilityPack.HtmlDocument]::new()
    $doc.LoadHtml($HtmlContent)

    $baseDate = [datetime]::ParseExact($PeriodStart, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
    $detectedDays = [datetime]::DaysInMonth($baseDate.Year, $baseDate.Month)

    # DetectHoursPerDay
    $hoursPerDay = 10
    $headerRows = $doc.DocumentNode.SelectNodes("//*[contains(@class, 'wsrow') and not(.//*[@data-user])]")
    if ($null -ne $headerRows) {
        foreach ($hRow in $headerRows) {
            $timeMatches = [regex]::Matches($hRow.InnerText, '\b(\d{2}):00\b')
            if ($timeMatches.Count -ge 2) {
                $firstHour = $timeMatches[0].Value
                for ($i = 1; $i -lt $timeMatches.Count; $i++) {
                    if ($timeMatches[$i].Value -eq $firstHour) {
                        $hoursPerDay = $i
                        break
                    }
                }
                break
            }
        }
    }

    # DetectStartHour
    $startHour = 8
    if ($null -ne $headerRows) {
        foreach ($hRow in $headerRows) {
            $m = [regex]::Match($hRow.InnerText, '\b(\d{2}):00\b')
            if ($m.Success) {
                $startHour = [int]$m.Groups[1].Value
                break
            }
        }
    }

    # DetectHourWidthPx
    $hourWidthPx = 3.0
    $widthSource = "FALLBACK (3px)"
    $widthDetected = $false

    if ($null -ne $headerRows) {
        foreach ($hRow in $headerRows) {
            $styledCells = $hRow.SelectNodes(".//*[contains(@style, 'width') and not(contains(@class, 'wscell'))]")
            if ($null -ne $styledCells -and $styledCells.Count -ge ($detectedDays - 2)) {
                $cellPx = Parse-PxValue (Extract-CssProperty -Style $styledCells[0].GetAttributeValue("style", "") -PropertyName "width")
                if ($null -ne $cellPx -and $cellPx -gt 0) {
                    if ($styledCells.Count -le ($detectedDays + 2)) {
                        $hourWidthPx = $cellPx / $hoursPerDay
                        $widthSource = "HEADER_DAY_CELL ($($cellPx)px / $($hoursPerDay)h)"
                    } else {
                        $hourWidthPx = $cellPx
                        $widthSource = "HEADER_HOUR_CELL ($($cellPx)px)"
                    }
                    $widthDetected = $true
                    break
                }
            }
        }
    }

    if (-not $widthDetected) {
        $firstUserRow = $doc.DocumentNode.SelectSingleNode("//*[contains(@class, 'wsrow') and .//*[@data-user]]")
        $sampleCells = if ($firstUserRow) { $firstUserRow.SelectNodes(".//*[contains(@class, 'calendar_cell') and @data-user]") } else { $null }
        if ($null -ne $sampleCells -and $sampleCells.Count -gt 0) {
            $px = Parse-PxValue (Extract-CssProperty -Style $sampleCells[0].GetAttributeValue("style", "") -PropertyName "width")
            if ($null -ne $px -and $px -gt 0) {
                if ($sampleCells.Count -le ($detectedDays + 2)) {
                    $hourWidthPx = $px / $hoursPerDay
                    $widthSource = "CALENDAR_DAY_CELL ($($px)px / $($hoursPerDay)h)"
                } else {
                    $hourWidthPx = $px
                    $widthSource = "CALENDAR_HOUR_CELL ($($px)px)"
                }
                $widthDetected = $true
            }
        }
    }

    if (-not $widthDetected) {
        $styleNodes = $doc.DocumentNode.SelectNodes("//style")
        if ($null -ne $styleNodes) {
            foreach ($sNode in $styleNodes) {
                $m = [regex]::Match($sNode.InnerText, '\.calendar_cell\s*\{[^}]*width\s*:\s*([\d.,]+)\s*px', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($m.Success) {
                    [double]$cssPx = 0
                    if ([double]::TryParse($m.Groups[1].Value.Replace(',', '.'), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$cssPx) -and $cssPx -gt 0) {
                        $hourWidthPx = $cssPx
                        $widthSource = "INLINE_STYLE ($($cssPx)px)"
                        break
                    }
                }
            }
        }
    }

    $usersList = [System.Collections.Generic.List[object]]::new()

    $rows = $doc.DocumentNode.SelectNodes("//*[contains(@class, 'wsrow')]")
    if ($null -ne $rows) {
        foreach ($row in $rows) {
            $userCells = $row.SelectNodes(".//*[contains(@class, 'calendar_cell') and @data-user]")
            if ($null -eq $userCells -or $userCells.Count -eq 0) { continue }

            $userId = $userCells[0].GetAttributeValue("data-user", "").Trim()
            $firstCell = $row.SelectSingleNode(".//*[contains(@class, 'wscell')]")
            $displayName = [HtmlAgilityPack.HtmlEntity]::DeEntitize($(if ($firstCell) { $firstCell.InnerText } else { "" })).Trim()

            $cellsPerDay = [Math]::Max(1, [int][Math]::Round([double]$userCells.Count / $detectedDays))

            $cellIndexMap = @{}
            for ($i = 0; $i -lt $userCells.Count; $i++) {
                $cellIndexMap[$userCells[$i]] = $i
            }

            $entriesList = [System.Collections.Generic.List[object]]::new()
            $elements = $row.SelectNodes(".//*[contains(@class, 'calendar_element') and not(contains(@class, 'calendar_element_'))]")

            if ($null -ne $elements) {
                foreach ($el in $elements) {
                    $header = $el.SelectSingleNode(".//*[contains(@class, 'calendar_element_header')]")
                    if ($null -eq $header) { continue }

                    $headerStyle   = $header.GetAttributeValue("style", "")
                    $statusColor   = Extract-CssProperty -Style $headerStyle -PropertyName "--bc"
                    $description   = [HtmlAgilityPack.HtmlEntity]::DeEntitize($header.GetAttributeValue("title", "")).Trim()
                    $onContextMenu = $header.GetAttributeValue("oncontextmenu", "")
                    $onDblClick    = $header.GetAttributeValue("ondblclick", "") + " " + $el.GetAttributeValue("ondblclick", "")

                    $leftPx  = Find-CssPxInHierarchy -StartNode $el -StopAtNode $row -PropertyName "left"
                    $widthPx = Find-CssPxInHierarchy -StartNode $el -StopAtNode $row -PropertyName "width"

                    $parentCell = $null
                    foreach ($anc in $el.Ancestors()) {
                        if ($cellIndexMap.ContainsKey($anc)) {
                            $parentCell = $anc
                            break
                        }
                    }

                    $cellIndex = if ($null -ne $parentCell) { $cellIndexMap[$parentCell] } else { 0 }
                    $cellDateAttr = if ($null -ne $parentCell) { $parentCell.GetAttributeValue("data-date", $null) } else { $null }

                    $leftVal = if ($null -ne $leftPx) { $leftPx } else { 0.0 }
                    if ($cellIndex -gt 0) {
                        $baseHoursFromCell = [double]$cellIndex * ($hoursPerDay / [double]$cellsPerDay)
                        $fineHoursFromLeft = $leftVal / $hourWidthPx
                        $totalHoursFromMonthStart = $baseHoursFromCell + $fineHoursFromLeft
                    } else {
                        $totalHoursFromMonthStart = $leftVal / $hourWidthPx
                    }

                    $dayOffset = [int][Math]::Floor($totalHoursFromMonthStart / $hoursPerDay)
                    $hourOffsetInDay = $totalHoursFromMonthStart - ($dayOffset * $hoursPerDay)
                    $startMinutesOffset = [Math]::Round($hourOffsetInDay * 60.0 / 15.0) * 15.0

                    $entryDate = $baseDate.AddDays([Math]::Max(0, $dayOffset))
                    $detectionSource = "GRID_PIXELS"

                    [datetime]$parsedCellDate = [datetime]::MinValue
                    if (-not [string]::IsNullOrWhiteSpace($cellDateAttr) -and
                        [datetime]::TryParseExact($cellDateAttr.Trim(), "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedCellDate)) {
                        $entryDate = $parsedCellDate
                        $detectionSource = "CELL_DATA_DATE"
                    }

                    foreach ($src in @($onDblClick, $onContextMenu, $description)) {
                        if ([string]::IsNullOrWhiteSpace($src)) { continue }
                        $isoM = [regex]::Match($src, '\b(20\d{2})-(\d{2})-(\d{2})\b')
                        if ($isoM.Success -and [datetime]::TryParseExact($isoM.Value, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedCellDate)) {
                            $entryDate = $parsedCellDate
                            $detectionSource += "+ATTR_DATE"
                            break
                        }
                        $deM = [regex]::Match($src, '\b(\d{2})\.(\d{2})\.(20\d{2})\b')
                        if ($deM.Success -and [datetime]::TryParseExact($deM.Value, "dd.MM.yyyy", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedCellDate)) {
                            $entryDate = $parsedCellDate
                            $detectionSource += "+ATTR_DATE"
                            break
                        }
                    }

                    $startDateTime = $entryDate.AddHours($startHour).AddMinutes($startMinutesOffset)
                    $durationHours = $null
                    $endDateTime = $null

                    if ($null -ne $widthPx -and $hourWidthPx -gt 0) {
                        $rawHours = $widthPx / $hourWidthPx
                        $totalMinutes = [Math]::Max(15.0, [Math]::Round($rawHours * 60.0 / 15.0) * 15.0)
                        $durationHours = [Math]::Round($totalMinutes / 60.0, 2)
                        $endDateTime = $startDateTime.AddMinutes($totalMinutes)
                    }

                    foreach ($src in @($onDblClick, $description)) {
                        if ([string]::IsNullOrWhiteSpace($src)) { continue }
                        $tMatch = [regex]::Match($src, '\b([01]?\d|2[0-3]):([0-5]\d)\s*(?:-|bis)\s*([01]?\d|2[0-3]):([0-5]\d)\b')
                        [timespan]$expStart = [timespan]::Zero
                        [timespan]$expEnd = [timespan]::Zero
                        if ($tMatch.Success -and
                            [timespan]::TryParse("$($tMatch.Groups[1].Value):$($tMatch.Groups[2].Value)", [ref]$expStart) -and
                            [timespan]::TryParse("$($tMatch.Groups[3].Value):$($tMatch.Groups[4].Value)", [ref]$expEnd) -and
                            $expEnd -gt $expStart) {

                            $startDateTime = $entryDate.Date.Add($expStart)
                            $endDateTime   = $entryDate.Date.Add($expEnd)
                            $durationHours = [Math]::Round(($expEnd - $expStart).TotalHours, 2)
                            $detectionSource += "+ATTR_TIME"
                            break
                        }
                    }

                    $labels = [System.Collections.Generic.List[string]]::new()
                    $leafDivs = $el.SelectNodes(".//div[not(.//div)]")
                    if ($null -ne $leafDivs) {
                        foreach ($d in $leafDivs) {
                            $t = [HtmlAgilityPack.HtmlEntity]::DeEntitize($d.InnerText).Trim()
                            if (-not [string]::IsNullOrWhiteSpace($t)) { $labels.Add($t) }
                        }
                    }

                    $executionType = "NONE"
                    if ($null -ne $el.SelectSingleNode(".//i[contains(@class, 'bi-car-front')]")) {
                        $executionType = "ON_SITE"
                    } elseif ($null -ne $el.SelectSingleNode(".//i[contains(@class, 'bi-pc-display-horizontal')]")) {
                        $executionType = "REMOTE"
                    }

                    $category = "UNKNOWN"
                    if (-not [string]::IsNullOrWhiteSpace($onContextMenu)) {
                        if ($onContextMenu -match "Absence")         { $category = "ABSENCE" }
                        elseif ($onContextMenu -match "Specialtime") { $category = "SPECIALTIME" }
                        elseif ($onContextMenu -match "Meeting")     { $category = "MEETING" }
                        else                                         { $category = "OTHER" }
                    }

                    $status = "UNKNOWN"
                    if (-not [string]::IsNullOrWhiteSpace($statusColor)) {
                        switch ($statusColor.ToUpperInvariant()) {
                            { $_ -in @("#91D9F0", "#A9F5A9") } { $status = "CONFIRMED"; break }
                            { $_ -in @("#F5B7B1", "#E6B0AA") } { $status = "CANCELLED_OR_ABSENT"; break }
                            "#F1F1AB"                          { $status = "BUFFER_OR_TENTATIVE"; break }
                            default                            { $status = "OTHER" }
                        }
                    }

                    $entriesList.Add([ordered]@{
                        entryId             = $header.GetAttributeValue("id", "").Trim()
                        category            = $category
                        description         = $description
                        date                = $entryDate.ToString("yyyy-MM-dd")
                        startDateTime       = $startDateTime.ToString("yyyy-MM-dd HH:mm")
                        endDateTime         = if ($endDateTime) { $endDateTime.ToString("yyyy-MM-dd HH:mm") } else { $null }
                        durationHours       = $durationHours
                        timeDetectionSource = $detectionSource
                        rawLabels           = $labels.ToArray()
                        executionType       = $executionType
                        status              = $status
                        statusColor         = $statusColor
                        layoutPosition      = [ordered]@{
                            cellIndex         = $cellIndex
                            cellDateAttribute = $cellDateAttr
                            leftPx            = $leftPx
                            widthPx           = $widthPx
                        }
                        meetingDetails      = $null
                        absenceDetails      = $null
                        specialtimeDetails  = $null
                    })
                }
            }

            $usersList.Add([ordered]@{
                userId      = $userId
                displayName = $displayName
                entryCount  = $entriesList.Count
                entries     = $entriesList.ToArray()
            })
        }
    }

    return [ordered]@{
        calendarMeta = [ordered]@{
            extractedAtUtc       = [datetime]::UtcNow.ToString("o")
            periodStart          = $PeriodStart
            hoursPerDay          = $hoursPerDay
            detectedDays         = $detectedDays
            detectedDayWidthPx   = [Math]::Round($hourWidthPx * $hoursPerDay, 2)
            detectedHourWidthPx  = [Math]::Round($hourWidthPx, 4)
            widthDetectionSource = $widthSource
            workingHours         = [ordered]@{
                start = "{0:D2}:00" -f $startHour
                end   = "{0:D2}:00" -f ($startHour + $hoursPerDay)
            }
        }
        users = $usersList.ToArray()
    }
}

function Set-ExactTimesFromMeeting {
    param($Entry, $Details)
    if ($null -eq $Details -or $null -eq $Details.timePeriod) { return }
    $tp = $Details.timePeriod
    if (-not [string]::IsNullOrWhiteSpace($tp.resolvedDateFrom)) {
        $Entry.date = $tp.resolvedDateFrom
        $Entry.timeDetectionSource = "MODAL_MEETING_DB_ID"
        if (-not [string]::IsNullOrWhiteSpace($tp.resolvedTimeFrom))  { $Entry.startDateTime = "$($tp.resolvedDateFrom) $($tp.resolvedTimeFrom)" }
        if (-not [string]::IsNullOrWhiteSpace($tp.resolvedTimeUntil)) { $Entry.endDateTime   = "$($tp.resolvedDateFrom) $($tp.resolvedTimeUntil)" }
        if ($null -ne $tp.durationHours) { $Entry.durationHours = $tp.durationHours }
    }
}

function Set-ExactTimesFromAbsence {
    param($Entry, $Details)
    if ($null -eq $Details -or [string]::IsNullOrWhiteSpace($Details.dateStart)) { return }
    $Entry.date = $Details.dateStart
    $Entry.timeDetectionSource = "MODAL_ABSENCE_DB_ID"

    $stopDate = if (-not [string]::IsNullOrWhiteSpace($Details.dateStop)) { $Details.dateStop } else { $Details.dateStart }
    $startTime = "08:00"
    $endTime   = "17:00"

    if ($Details.absenceTypeId -eq "1") {
        $startTime = "08:00"
        $endTime   = "12:30"
    } elseif ($Details.absenceTypeId -eq "2") {
        $startTime = "12:30"
        $endTime   = "17:00"
    }

    $Entry.startDateTime = "$($Details.dateStart) $startTime"
    $Entry.endDateTime   = "$stopDate $endTime"

    [datetime]$dStart = [datetime]::MinValue
    [datetime]$dStop  = [datetime]::MinValue
    if ([datetime]::TryParseExact($Details.dateStart, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dStart) -and
        [datetime]::TryParseExact($stopDate, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dStop) -and
        $dStop -ge $dStart) {
        $days = ($dStop - $dStart).Days + 1
        $hoursPerDay = if ($Details.absenceTypeId -in @("1", "2")) { 4.0 } else { 8.0 }
        $Entry.durationHours = $days * $hoursPerDay
    }
}

function Set-ExactTimesFromSpecialtime {
    param($Entry, $Details)
    if ($null -eq $Details -or $null -eq $Details.timePeriod) { return }
    $tp = $Details.timePeriod
    if (-not [string]::IsNullOrWhiteSpace($tp.resolvedDateFrom)) {
        $Entry.date = $tp.resolvedDateFrom
        $Entry.timeDetectionSource = "MODAL_SPECIALTIME_DB_ID"
        if (-not [string]::IsNullOrWhiteSpace($tp.resolvedTimeFrom))  { $Entry.startDateTime = "$($tp.resolvedDateFrom) $($tp.resolvedTimeFrom)" }
        if (-not [string]::IsNullOrWhiteSpace($tp.resolvedTimeUntil)) { $Entry.endDateTime   = "$($tp.resolvedDateFrom) $($tp.resolvedTimeUntil)" }
        if ($null -ne $tp.durationHours) { $Entry.durationHours = $tp.durationHours }
    }
}

# ============================================================================
# 5. Benutzerauswahl & Filter (komplett über credentials.txt)
# ============================================================================
function Resolve-ManualUserSelection {
    param($Users, [string]$InputStr)
    $selectedIndices = [System.Collections.Generic.HashSet[int]]::new()
    $tokens = $InputStr -split '[,;]'

    foreach ($rawToken in $tokens) {
        $token = $rawToken.Trim()
        if ([string]::IsNullOrEmpty($token)) { continue }

        $rangeMatch = [regex]::Match($token, '^(\d+)\s*-\s*(\d+)$')
        if ($rangeMatch.Success) {
            $sIdx = [int]$rangeMatch.Groups[1].Value
            $eIdx = [int]$rangeMatch.Groups[2].Value
            $from = [Math]::Max(1, [Math]::Min($sIdx, $eIdx))
            $to   = [Math]::Min($Users.Count, [Math]::Max($sIdx, $eIdx))
            for ($idx = $from; $idx -le $to; $idx++) {
                $selectedIndices.Add($idx - 1) | Out-Null
            }
            continue
        }

        [int]$singleIdx = 0
        if ([int]::TryParse($token, [ref]$singleIdx)) {
            if ($singleIdx -ge 1 -and $singleIdx -le $Users.Count) {
                $selectedIndices.Add($singleIdx - 1) | Out-Null
            }
            continue
        }

        for ($i = 0; $i -lt $Users.Count; $i++) {
            if ($Users[$i].userId.Equals($token, [StringComparison]::OrdinalIgnoreCase) -or
                $Users[$i].displayName.IndexOf($token, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $selectedIndices.Add($i) | Out-Null
            }
        }
    }

    return $selectedIndices | Sort-Object | ForEach-Object { $Users[$_] }
}

function Prompt-UserSelectionFilter {
    param($Users)
    $selectedCount = ($Users | Where-Object { $_.isSelected }).Count

    Write-Host "`nWelche Benutzer sollen aus DF abgerufen werden?"
    Write-Host "  [1] Alle Benutzer ($($Users.Count))"
    Write-Host "  [2] Nur in DF vorausgewählte Benutzer ($selectedCount)"
    Write-Host "  [3] Bestimmte Benutzer auswählen"

    while ($true) {
        $choice = (Read-Host "Auswahl [Enter = 1 (Alle)]").Trim()
        if ([string]::IsNullOrEmpty($choice) -or $choice -eq "1") { return "ALL" }
        if ($choice -eq "2") { return "SELECTED" }

        if ($choice -eq "3") {
            Write-Host "`nVerfügbare Benutzer ([x] = in DF ausgewählt):"
            for ($i = 0; $i -lt $Users.Count; $i++) {
                $mark = if ($Users[$i].isSelected) { "x" } else { " " }
                Write-Host ("  {0,2}. [{1}] {2}" -f ($i + 1), $mark, $Users[$i].displayName)
            }

            while ($true) {
                $rawSelection = (Read-Host "`nNummern (z. B. 1, 3, 5-8) oder Namensteil(e) (z. B. Boeshaar, Müller) eingeben").Trim()
                if ([string]::IsNullOrWhiteSpace($rawSelection)) { continue }

                $matched = @(Resolve-ManualUserSelection -Users $Users -InputStr $rawSelection)
                if ($matched.Count -gt 0) {
                    $names = ($matched | ForEach-Object { $_.displayName }) -join ", "
                    Write-Host "Ausgewählt ($($matched.Count)): $names"
                    return (($matched | ForEach-Object { $_.userId }) -join ",")
                }
                Write-Host "Keine passenden Benutzer gefunden. Bitte erneut eingeben." -ForegroundColor Yellow
            }
        }
        Write-Host "Ungültige Eingabe. Bitte 1, 2 oder 3 eingeben." -ForegroundColor Yellow
    }
}

function Apply-UserFilter {
    param($Users, [string]$UserFilter)
    if ([string]::IsNullOrWhiteSpace($UserFilter) -or $UserFilter.Equals("ALL", [StringComparison]::OrdinalIgnoreCase)) {
        return $Users
    }
    if ($UserFilter.Equals("SELECTED", [StringComparison]::OrdinalIgnoreCase)) {
        return @($Users | Where-Object { $_.isSelected })
    }

    $allowedIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($id in ($UserFilter -split '[,;]')) {
        if (-not [string]::IsNullOrWhiteSpace($id)) { $allowedIds.Add($id.Trim()) | Out-Null }
    }
    return @($Users | Where-Object { $allowedIds.Contains($_.userId) })
}

function Get-OrFetchUsersJson {
    param($Session, [hashtable]$Config, [string]$CredentialsPath)

    Write-Host "Rufe Benutzerliste von 'dropdown_selection_user.php' ab..."
    $dropdownHtml = Get-DFUsersDropdownHtml -Session $Session
    $availableUsers = @()

    if (-not [string]::IsNullOrWhiteSpace($dropdownHtml)) {
        $fetchedUsers = @(Parse-UserDropdown -Html $dropdownHtml)
        if ($fetchedUsers.Count -gt 0) {
            $availableUsers = $fetchedUsers
            $Config.CachedUsers = $fetchedUsers

            $selectedCount  = ($fetchedUsers | Where-Object { $_.isSelected }).Count
            $availableCount = $fetchedUsers.Count - $selectedCount
            Write-Host "$($fetchedUsers.Count) User gefunden ($selectedCount in DF ausgewählt, $availableCount verfügbar)."

            if ([string]::IsNullOrWhiteSpace($Config.UserFilter)) {
                $Config.UserFilter = Prompt-UserSelectionFilter -Users $fetchedUsers

                if ($Config.PersistInFile -or (Test-Path $CredentialsPath)) {
                    $saveChoice = (Read-Host "Benutzerauswahl in 'credentials.txt' speichern? (j/n) [j]").Trim()
                    if ([string]::IsNullOrEmpty($saveChoice) -or $saveChoice -match '^(j|y)$') {
                        Save-AppConfig -FilePath $CredentialsPath -Config $Config
                        Write-Host "Benutzerauswahl & User-Liste in 'credentials.txt' gespeichert.`n"
                    }
                }
            }
            elseif ($Config.PersistInFile -or (Test-Path $CredentialsPath)) {
                Save-AppConfig -FilePath $CredentialsPath -Config $Config
            }
        }
    }

    if ($availableUsers.Count -eq 0 -and $null -ne $Config.CachedUsers -and $Config.CachedUsers.Count -gt 0) {
        Write-Host "Fallback: Lade $($Config.CachedUsers.Count) Benutzer aus 'CachedUsers' in 'credentials.txt'." -ForegroundColor Yellow
        $availableUsers = @($Config.CachedUsers)
    }

    if ($availableUsers.Count -gt 0) {
        $activeUsers = @(Apply-UserFilter -Users $availableUsers -UserFilter $Config.UserFilter)
        if ($activeUsers.Count -eq 0) {
            Write-Host "Warnung: Der gespeicherte UserFilter ergab 0 Treffer. Lade stattdessen alle Benutzer." -ForegroundColor Yellow
            $activeUsers = $availableUsers
        }

        $filteredIds = @($activeUsers | ForEach-Object { $_.userId } | Select-Object -Unique)
        $compactJson = ConvertTo-Json -InputObject @($filteredIds) -Compress

        Write-Host "Aktive Benutzer für Kalender-Abruf ($($filteredIds.Count)):"
        foreach ($u in $activeUsers) {
            Write-Host "  - $($u.displayName) ($($u.userId))"
        }
        Write-Host ""

        return @{
            CompactUsersJson = $compactJson
            AllUsers         = $availableUsers
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Config.UserFilter) -and
        $Config.UserFilter -notmatch '^(ALL|SELECTED)$') {

        $directIds = @($Config.UserFilter -split '[,;]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique)
        if ($directIds.Count -gt 0) {
            Write-Host "Fallback: Verwende $($directIds.Count) direkte User-IDs aus 'UserFilter' in 'credentials.txt'.`n" -ForegroundColor Yellow
            return @{
                CompactUsersJson = (ConvertTo-Json -InputObject @($directIds) -Compress)
                AllUsers         = @()
            }
        }
    }

    return @{ CompactUsersJson = $null; AllUsers = @() }
}

# ============================================================================
# 6. Lokale Outlook-Synchronisation (COM)
# ============================================================================
function Get-OutlookSubCalendarNames {
    $names = [System.Collections.Generic.List[string]]::new()
    $outlookApp = $null
    $ns = $null
    $defaultCalendar = $null

    try {
        $outlookApp = New-Object -ComObject Outlook.Application
        $ns = $outlookApp.GetNamespace("MAPI")
        $ns.Logon("", "", $false, $false)
        $defaultCalendar = $ns.GetDefaultFolder(9) # olFolderCalendar = 9

        foreach ($folder in $defaultCalendar.Folders) {
            try {
                if ([int]$folder.DefaultItemType -eq 1) { # olAppointmentItem = 1
                    $names.Add([string]$folder.Name)
                }
            }
            finally {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null
            }
        }
    }
    catch {
        # Falls Outlook nicht bereit ist, leere Liste zurückgeben
    }
    finally {
        if ($defaultCalendar) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($defaultCalendar) | Out-Null }
        if ($ns)              { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns) | Out-Null }
        if ($outlookApp)      { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlookApp) | Out-Null }
    }

    return $names
}

function Find-BestUserMatchIndex {
    param($Users, [string]$LoginUsername)
    if ([string]::IsNullOrWhiteSpace($LoginUsername)) { return -1 }

    $cleanLogin = $LoginUsername.Split('@')[0]
    $parts = @($cleanLogin -split '[._\-\s]+' | Where-Object { $_.Length -ge 2 })
    if ($parts.Count -eq 0) { return -1 }

    $bestIdx = -1
    $bestScore = 0

    for ($i = 0; $i -lt $Users.Count; $i++) {
        $name = if ($Users[$i].displayName) { $Users[$i].displayName } else { "" }
        $score = ($parts | Where-Object { $name.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -ge 0 }).Count
        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestIdx = $i
        }
    }

    return $(if ($bestScore -gt 0) { $bestIdx } else { -1 })
}

function Resolve-OrPromptOwnUserId {
    param($Users, [string]$LoginUsername)
    if ($null -eq $Users -or $Users.Count -eq 0) { return $null }

    if ($Users.Count -eq 1) {
        Write-Host "`nEigener DF-Kalender automatisch erkannt: $($Users[0].displayName)"
        return $Users[0].userId
    }

    $suggestedIdx = Find-BestUserMatchIndex -Users $Users -LoginUsername $LoginUsername

    Write-Host "`nWelcher DF-Benutzer bist DU (für die Übertragung in dein eigenes Outlook)?"
    for ($i = 0; $i -lt $Users.Count; $i++) {
        $hint = if ($i -eq $suggestedIdx) { "  <-- Vorschlag" } else { "" }
        Write-Host ("  {0,2}. {1}{2}" -f ($i + 1), $Users[$i].displayName, $hint)
    }

    while ($true) {
        $prompt = if ($suggestedIdx -ge 0) {
            "Deine Nummer [Enter = $($suggestedIdx + 1) ($($Users[$suggestedIdx].displayName))]: "
        } else {
            "Deine Nummer eingeben: "
        }

        Write-Host -NoNewline $prompt
        $inputStr = (Read-Host).Trim()

        if ([string]::IsNullOrEmpty($inputStr) -and $suggestedIdx -ge 0) {
            return $Users[$suggestedIdx].userId
        }

        [int]$num = 0
        if ([int]::TryParse($inputStr, [ref]$num) -and $num -ge 1 -and $num -le $Users.Count) {
            $chosen = $Users[$num - 1]
            Write-Host "Eigener Kalender festgelegt auf: $($chosen.displayName)"
            return $chosen.userId
        }

        $nameMatches = @($Users | Where-Object { -not [string]::IsNullOrWhiteSpace($inputStr) -and $_.displayName.IndexOf($inputStr, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
        if ($nameMatches.Count -eq 1) {
            Write-Host "Eigener Kalender festgelegt auf: $($nameMatches[0].displayName)"
            return $nameMatches[0].userId
        }

        Write-Host "Ungültige Auswahl. Bitte eine gültige Nummer aus der Liste eingeben." -ForegroundColor Yellow
    }
}

function Prompt-OutlookCalendarSelection {
    $existingSubCalendars = Get-OutlookSubCalendarNames

    Write-Host "`nIn welchen lokalen Outlook-Kalender sollen deine Termine geschrieben werden?"
    Write-Host "  [1] Standard-Kalender (dein persönlicher Outlook-Hauptkalender)"

    $optionNumber = 2
    $optionMap = @{ 1 = "DEFAULT" }

    foreach ($folderName in $existingSubCalendars) {
        Write-Host "  [$optionNumber] Vorhandener Unterkalender '$folderName'"
        $optionMap[$optionNumber] = $folderName
        $optionNumber++
    }

    $hasDFFolder = $existingSubCalendars | Where-Object { $_.Equals("DF", [StringComparison]::OrdinalIgnoreCase) }
    if (-not $hasDFFolder) {
        Write-Host "  [$optionNumber] Neuen Unterkalender 'DF' anlegen"
        $optionMap[$optionNumber] = "DF"
        $optionNumber++
    }

    $customOption = $optionNumber++
    Write-Host "  [$customOption] Eigenen Unterkalender-Namen eingeben"

    $skipOption = $optionNumber
    Write-Host "  [$skipOption] Nicht mit Outlook synchronisieren"

    while ($true) {
        $choice = (Read-Host "Auswahl [Enter = 1 (Standard-Kalender)]").Trim()
        if ([string]::IsNullOrEmpty($choice)) { return "DEFAULT" }

        [int]$num = 0
        if ([int]::TryParse($choice, [ref]$num)) {
            if ($optionMap.ContainsKey($num)) { return $optionMap[$num] }

            if ($num -eq $customOption) {
                $customName = (Read-Host "Name des Outlook-Unterkalenders eingeben").Trim()
                if (-not [string]::IsNullOrWhiteSpace($customName)) { return $customName }
            }

            if ($num -eq $skipOption) { return "NONE" }
        }

        Write-Host "Ungültige Eingabe. Bitte eine der angezeigten Nummern wählen." -ForegroundColor Yellow
    }
}

function Build-OutlookSubject {
    param($Entry)
    $cancelledPrefix = if ($Entry.status -eq "CANCELLED_OR_ABSENT") { "[Abgesagt] " } else { "" }

    if ($Entry.category -eq "MEETING") {
        $customer  = if ($Entry.meetingDetails -and $Entry.meetingDetails.customerLines.Count -gt 0) { $Entry.meetingDetails.customerLines[0] } else { $null }
        $operation = if ($Entry.meetingDetails) { $Entry.meetingDetails.operationName } else { $null }
        $core      = if (-not [string]::IsNullOrWhiteSpace($customer)) { $customer } else { $Entry.description }

        if (-not [string]::IsNullOrWhiteSpace($operation) -and -not [string]::IsNullOrWhiteSpace($core)) {
            return "$cancelledPrefix$core ($operation)"
        }
        return "$cancelledPrefix$(if ($core) { $core } else { 'DF Termin' })"
    }

    if ($Entry.category -eq "ABSENCE") {
        $reason = if ($Entry.absenceDetails) { $Entry.absenceDetails.absenceReasonName } else { $null }
        $halfDay = switch ($Entry.absenceDetails.absenceTypeId) {
            "1"     { " (1. Halbtag)" }
            "2"     { " (2. Halbtag)" }
            default { "" }
        }
        $core = if (-not [string]::IsNullOrWhiteSpace($reason)) { $reason } else { $Entry.description }
        return "Fehlzeit: $core$halfDay"
    }

    if ($Entry.category -eq "SPECIALTIME") {
        $stType = if ($Entry.specialtimeDetails) { $Entry.specialtimeDetails.specialtimeTypeName } else { $null }
        $core   = if (-not [string]::IsNullOrWhiteSpace($stType)) { $stType } else { $Entry.description }
        return "Sonderzeit: $core"
    }

    return "$cancelledPrefix$($Entry.description)"
}

function Build-OutlookLocation {
    param($Entry)
    if ($Entry.category -eq "MEETING" -and $Entry.meetingDetails) {
        $lines = @($Entry.meetingDetails.customerLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($lines.Count -gt 1) {
            return ($lines[1..($lines.Count - 1)] -join ", ")
        }
        if ($Entry.executionType -eq "ON_SITE") { return "Vor Ort" }
        if ($Entry.executionType -eq "REMOTE")  { return "Remote" }
    }
    return ""
}

function Build-OutlookBody {
    param($User, $Entry)
    $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("Mitarbeiter: $($User.displayName)") | Out-Null
    $sb.AppendLine("Kategorie:   $($Entry.category)") | Out-Null

    if ($Entry.category -eq "MEETING" -and $Entry.meetingDetails) {
        $m = $Entry.meetingDetails
        if ($m.customerLines.Count -gt 0)    { $sb.AppendLine("Kunde:       $($m.customerLines -join ' | ')") | Out-Null }
        if ($m.contactLines.Count -gt 0)     { $sb.AppendLine("Kontakt:     $($m.contactLines -join ' | ')") | Out-Null }
        if ($m.orderLines.Count -gt 0)       { $sb.AppendLine("Auftrag:     $($m.orderLines -join ' | ')") | Out-Null }
        if ($m.ticket)                       { $sb.AppendLine("Ticket:      $($m.ticket)") | Out-Null }
        if ($m.operationName)                { $sb.AppendLine("Tätigkeit:   $($m.operationName)") | Out-Null }
        if ($m.useName)                      { $sb.AppendLine("Einsatzart:  $($m.useName)") | Out-Null }

        if (-not [string]::IsNullOrWhiteSpace($m.description)) {
            $sb.AppendLine("`n--- Beschreibung ---") | Out-Null
            $sb.AppendLine($m.description) | Out-Null
        }

        if ($m.history.Count -gt 0) {
            $sb.AppendLine("`n--- Historie ---") | Out-Null
            foreach ($h in $m.history) {
                $sb.AppendLine("[$($h.timestampRaw)] $($h.author): $($h.status)") | Out-Null
                if (-not [string]::IsNullOrWhiteSpace($h.description)) {
                    $sb.AppendLine($h.description) | Out-Null
                }
                $sb.AppendLine() | Out-Null
            }
        }
    }
    elseif ($Entry.category -eq "ABSENCE" -and $Entry.absenceDetails) {
        $a = $Entry.absenceDetails
        $sb.AppendLine("Zeitraum:    $($a.dateStart) bis $($a.dateStop)") | Out-Null
        if ($a.absenceTypeName)   { $sb.AppendLine("Typ:         $($a.absenceTypeName)") | Out-Null }
        if ($a.absenceReasonName) { $sb.AppendLine("Grund:       $($a.absenceReasonName)") | Out-Null }
        if (-not [string]::IsNullOrWhiteSpace($a.description)) {
            $sb.AppendLine("`n--- Beschreibung ---") | Out-Null
            $sb.AppendLine($a.description) | Out-Null
        }
    }
    elseif ($Entry.category -eq "SPECIALTIME" -and $Entry.specialtimeDetails) {
        $s = $Entry.specialtimeDetails
        if ($s.specialtimeTypeName) { $sb.AppendLine("Sonderzeit:  $($s.specialtimeTypeName)") | Out-Null }
        if (-not [string]::IsNullOrWhiteSpace($s.description)) {
            $sb.AppendLine("`n--- Beschreibung ---") | Out-Null
            $sb.AppendLine($s.description) | Out-Null
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($Entry.description)) {
        $sb.AppendLine("`n--- Beschreibung ---") | Out-Null
        $sb.AppendLine($Entry.description) | Out-Null
    }

    $sb.AppendLine("`n[DF DB_ID: $($Entry.entryId)]") | Out-Null
    return $sb.ToString()
}

function Sync-OutlookOwnEntries {
    param($Exports, [string]$OwnUserId, [string]$CalendarSetting)

    $outlookApp = $null
    $ns = $null
    $defaultCalendar = $null
    $targetFolder = $null

    try {
        $outlookApp = New-Object -ComObject Outlook.Application
        $ns = $outlookApp.GetNamespace("MAPI")
        $ns.Logon("", "", $false, $false)

        $defaultCalendar = $ns.GetDefaultFolder(9) # olFolderCalendar

        if ([string]::IsNullOrWhiteSpace($CalendarSetting) -or $CalendarSetting.Equals("DEFAULT", [StringComparison]::OrdinalIgnoreCase)) {
            $targetFolder = $defaultCalendar
            Write-Host "`nSynchronisiere eigene Termine in Outlook Standard-Kalender..."
        }
        else {
            foreach ($folder in $defaultCalendar.Folders) {
                if ([string]::Equals([string]$folder.Name, $CalendarSetting, [StringComparison]::OrdinalIgnoreCase)) {
                    $targetFolder = $folder
                    break
                }
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null
            }
            if ($null -eq $targetFolder) {
                $targetFolder = $defaultCalendar.Folders.Add($CalendarSetting, 9)
            }
            Write-Host "`nSynchronisiere eigene Termine in Outlook Unterkalender '$CalendarSetting'..."
        }

        $existingByKey = @{}
        $items = $targetFolder.Items
        $count = $items.Count
        for ($i = 1; $i -le $count; $i++) {
            $item = $items.Item($i)
            try {
                $prop = $item.UserProperties.Find("DFKey")
                if ($null -ne $prop -and $null -ne $prop.Value) {
                    $key = $prop.Value.ToString()
                    if (-not [string]::IsNullOrWhiteSpace($key) -and -not $existingByKey.ContainsKey($key)) {
                        $existingByKey[$key] = $item
                    }
                }
            }
            catch {
                # Nicht-Termin-Elemente ignorieren
            }
        }

        $processedKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $createdCount = 0
        $updatedCount = 0
        $skippedCount = 0

        foreach ($export in $Exports) {
            $ownUsers = @($export.users | Where-Object { $_.userId.Equals($OwnUserId, [StringComparison]::OrdinalIgnoreCase) })

            foreach ($user in $ownUsers) {
                foreach ($entry in $user.entries) {
                    if ([string]::IsNullOrWhiteSpace($entry.entryId)) {
                        $skippedCount++
                        continue
                    }

                    $syncKey = "$($user.userId)_$($entry.entryId)"
                    if (-not $processedKeys.Add($syncKey)) { continue }

                    $isAllDay = $false
                    [datetime]$startDt = [datetime]::MinValue
                    [datetime]$endDt   = [datetime]::MinValue
                    $rangeValid = $false

                    if ($entry.category -eq "ABSENCE" -and $entry.absenceDetails) {
                        $abs = $entry.absenceDetails
                        [datetime]$dStart = [datetime]::MinValue
                        if ([datetime]::TryParseExact($abs.dateStart, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dStart)) {
                            $dStop = $dStart
                            if (-not [string]::IsNullOrWhiteSpace($abs.dateStop)) {
                                [datetime]::TryParseExact($abs.dateStop, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dStop) | Out-Null
                            }
                            if ($dStop -lt $dStart) { $dStop = $dStart }

                            if ($abs.absenceTypeId -eq "0" -or [string]::IsNullOrWhiteSpace($abs.absenceTypeId)) {
                                $isAllDay   = $true
                                $startDt    = $dStart.Date
                                $endDt      = $dStop.Date.AddDays(1)
                                $rangeValid = $true
                            }
                        }
                    }

                    if (-not $rangeValid) {
                        if (-not [datetime]::TryParseExact($entry.startDateTime, "yyyy-MM-dd HH:mm", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$startDt)) {
                            [datetime]$fbDate = [datetime]::MinValue
                            if (-not [datetime]::TryParseExact($entry.date, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$fbDate)) {
                                $skippedCount++
                                continue
                            }
                            $startDt = $fbDate.AddHours(8)
                        }

                        if (-not [datetime]::TryParseExact($entry.endDateTime, "yyyy-MM-dd HH:mm", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$endDt) -or $endDt -le $startDt) {
                            $hours = if ($null -ne $entry.durationHours) { [double]$entry.durationHours } else { 1.0 }
                            $endDt = $startDt.AddHours([Math]::Max(0.25, $hours))
                        }
                    }

                    $appt = $null
                    $isUpdate = $false

                    if ($existingByKey.ContainsKey($syncKey)) {
                        $appt = $existingByKey[$syncKey]
                        $isUpdate = $true
                    }
                    else {
                        $appt = $targetFolder.Items.Add(1) # olAppointmentItem
                        $prop = $appt.UserProperties.Add("DFKey", 1, $false) # olText = 1
                        $prop.Value = $syncKey
                    }

                    $appt.Subject     = Build-OutlookSubject -Entry $entry
                    $appt.AllDayEvent = $isAllDay
                    $appt.Start       = $startDt
                    $appt.End         = $endDt
                    $appt.Location    = Build-OutlookLocation -Entry $entry
                    $appt.Body        = Build-OutlookBody -User $user -Entry $entry
                    $appt.Categories  = switch ($entry.category) {
                        "MEETING"     { "DF; Termin" }
                        "ABSENCE"     { "DF; Fehlzeit" }
                        "SPECIALTIME" { "DF; Sonderzeit" }
                        default       { "DF" }
                    }
                    $appt.BusyStatus  = if ($entry.category -eq "ABSENCE") { 3 } # olOutOfOffice
                                        elseif ($entry.status -eq "CANCELLED_OR_ABSENT") { 0 } # olFree
                                        elseif ($entry.status -eq "BUFFER_OR_TENTATIVE") { 1 } # olTentative
                                        else { 2 } # olBusy
                    $appt.ReminderSet = $false

                    $appt.Save()

                    if ($isUpdate) { $updatedCount++ } else { $createdCount++ }
                }
            }
        }

        Write-Host "Outlook-Sync abgeschlossen: $createdCount neu erstellt, $updatedCount aktualisiert, $skippedCount übersprungen." -ForegroundColor Green
    }
    finally {
        if ($targetFolder -and ($targetFolder -ne $defaultCalendar)) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($targetFolder) | Out-Null
        }
        if ($defaultCalendar) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($defaultCalendar) | Out-Null }
        if ($ns)              { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns) | Out-Null }
        if ($outlookApp)      { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlookApp) | Out-Null }
    }
}

function Invoke-ConfigureAndSyncOutlook {
    param($Exports, $AvailableUsers, [hashtable]$Config, [string]$CredentialsPath)

    if ([string]::Equals($Config.OutlookCalendar, "NONE", [StringComparison]::OrdinalIgnoreCase)) {
        Write-Host "Outlook-Sync ist in 'credentials.txt' deaktiviert (OutlookCalendar=NONE)."
        return
    }

    $configChanged = $false

    if ([string]::IsNullOrWhiteSpace($Config.OwnUserId)) {
        if ($null -eq $AvailableUsers -or $AvailableUsers.Count -eq 0) {
            $AvailableUsers = @(
                $Exports | ForEach-Object { $_.users } | Group-Object -Property userId | ForEach-Object {
                    [ordered]@{
                        userId      = $_.Name
                        displayName = $_.Group[0].displayName
                    }
                }
            )
        }

        $Config.OwnUserId = Resolve-OrPromptOwnUserId -Users $AvailableUsers -LoginUsername $Config.Username
        if (-not [string]::IsNullOrWhiteSpace($Config.OwnUserId)) {
            $configChanged = $true
        }
    }

    if ([string]::IsNullOrWhiteSpace($Config.OwnUserId)) {
        Write-Host "Kein eigener DF-Benutzer ausgewählt - Outlook-Sync wird übersprungen."
        return
    }

    $ownSchedules = @($Exports | ForEach-Object { $_.users } | Where-Object { $_.userId.Equals($Config.OwnUserId, [StringComparison]::OrdinalIgnoreCase) })
    if ($ownSchedules.Count -eq 0) {
        Write-Host "Hinweis: Dein eigener Benutzer ($($Config.OwnUserId)) war im geladenen DF-Abruf nicht enthalten." -ForegroundColor Yellow
        Write-Host "Tipp: Prüfe 'UserFilter' in der 'credentials.txt', damit dein eigener Kalender mit abgerufen wird."
        return
    }

    $ownEntryCount  = ($ownSchedules | ForEach-Object { $_.entries.Count } | Measure-Object -Sum).Sum
    $ownDisplayName = $ownSchedules[0].displayName

    try {
        if ([string]::IsNullOrWhiteSpace($Config.OutlookCalendar)) {
            Write-Host "`nFür '$ownDisplayName' wurden $ownEntryCount eigene Einträge gefunden."
            $Config.OutlookCalendar = Prompt-OutlookCalendarSelection
            if (-not [string]::IsNullOrWhiteSpace($Config.OutlookCalendar)) {
                $configChanged = $true
            }
        }

        if ($configChanged -and ($Config.PersistInFile -or (Test-Path $CredentialsPath))) {
            Save-AppConfig -FilePath $CredentialsPath -Config $Config
            Write-Host "Outlook-Einstellungen in 'credentials.txt' gespeichert."
        }

        if ([string]::Equals($Config.OutlookCalendar, "NONE", [StringComparison]::OrdinalIgnoreCase)) {
            Write-Host "Outlook-Sync übersprungen."
            return
        }

        Sync-OutlookOwnEntries -Exports $Exports -OwnUserId $Config.OwnUserId -CalendarSetting $Config.OutlookCalendar
    }
    catch {
        Write-Host "`nFehler bei der Outlook-Synchronisation: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Hinweis: Stelle sicher, dass das klassische Outlook auf diesem Windows-PC installiert ist."
    }
}

# ============================================================================
# 7. Hauptprogramm (Main Execution)
# ============================================================================
function Invoke-Main {
    $credentialsPath = Join-Path $BaseDir "credentials.txt"

    # 1. URL, Workspace-ID, Zugangsdaten & gespeicherte Einstellungen laden / abfragen
    $config = Get-OrPromptAppConfig -FilePath $credentialsPath
    if ([string]::IsNullOrWhiteSpace($config.Username) -or [string]::IsNullOrWhiteSpace($config.Password)) {
        Write-Host "Abbruch: Keine Zugangsdaten eingegeben." -ForegroundColor Red
        return
    }

    # 2. Vorab prüfen, ob der Host eine 10.x.x.x VPN-IP hat
    if (-not (Test-DFVpnIp -BaseUrl $config.BaseUrl)) {
        Write-Host "Abbruch: Keine Verbindung zum internen VPN-Netz (10.x.x.x)." -ForegroundColor Red
        return
    }

    # 3. Zeitraum aus credentials.txt laden oder interaktiv abfragen
    $monthsToFetch = Get-OrPromptMonthsToFetch -Config $config -CredentialsPath $credentialsPath

    $allFetchedMonths  = [System.Collections.Generic.List[object]]::new()
    $allAvailableUsers = @()

    $session = New-DFSession -BaseUrl $config.BaseUrl
    try {
        Write-Host "`nVerbinde mit '$($config.BaseUrl)' und logge ein als '$($config.Username)'..."
        $loggedIn = Connect-DF -Session $session -Username $config.Username -Password $config.Password -WorkspaceId $config.WorkspaceId

        if (-not $loggedIn) {
            Write-Host "Login fehlgeschlagen! Bitte URL, Workspace-ID, Benutzername und Passwort prüfen." -ForegroundColor Red
            return
        }

        $sessId = Get-DFPhpSessionId -Session $session
        Write-Host "Login erfolgreich! PHPSESSID: $sessId`n" -ForegroundColor Green

        # 4. Benutzerliste abrufen & in credentials.txt cachen
        $userResult = Get-OrFetchUsersJson -Session $session -Config $config -CredentialsPath $credentialsPath
        $usersJson = $userResult.CompactUsersJson
        $allAvailableUsers = $userResult.AllUsers

        if ([string]::IsNullOrEmpty($usersJson)) {
            Write-Host "Keine gültigen User-IDs gefunden!" -ForegroundColor Red
            return
        }

        #$desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
        $desktopPath = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

        foreach ($monthDate in $monthsToFetch) {
            $periodStart = $monthDate.ToString("yyyy-MM-dd")
            $monthTag    = $monthDate.ToString("yyyy-MM")

            Write-Host "=================================================="
            Write-Host "Rufe Kalender für Monat $monthTag ($periodStart) ab..."
            Write-Host "==================================================`n"

            $calendarHtml = Get-DFCalendarMonthHtml -Session $session -DateStamp $periodStart -UsersJsonArray $usersJson -ScrollValue 229
            if ($null -eq $calendarHtml) {
                Write-Host "Überspringe Monat $monthTag wegen Abruffehler.`n" -ForegroundColor Yellow
                continue
            }

            $parsedCalendar = Parse-CalendarHtml -HtmlContent $calendarHtml -PeriodStart $periodStart

            $allEntries = [System.Collections.Generic.List[object]]::new()
            foreach ($u in $parsedCalendar.users) {
                foreach ($e in $u.entries) {
                    if (-not [string]::IsNullOrWhiteSpace($e.entryId) -and $e.category -in @("MEETING", "ABSENCE", "SPECIALTIME")) {
                        $allEntries.Add($e)
                    }
                }
            }

            Write-Host "[$monthTag] Rufe Detail-Modals (Mode=4) für $($allEntries.Count) Einträge ab...`n"

            $counter = 0
            foreach ($entry in $allEntries) {
                $counter++
                $entryId = $entry.entryId

                Write-Host -NoNewline "[$monthTag] [$counter/$($allEntries.Count)] [$($entry.category)] DB_ID=$entryId ... "

                $modalHtml = Get-DFModalHtml -Session $session -Category $entry.category -DbId $entryId

                if (-not [string]::IsNullOrWhiteSpace($modalHtml)) {
                    if ($entry.category -eq "MEETING") {
                        $details = Parse-MeetingModal -ModalHtml $modalHtml -FallbackEntry $entry
                        $entry.meetingDetails = $details
                        Set-ExactTimesFromMeeting -Entry $entry -Details $details

                        Start-Sleep -Milliseconds 50
                        $historyHtml = Get-DFMeetingHistoryHtml -Session $session -DbId $entryId
                        if (-not [string]::IsNullOrWhiteSpace($historyHtml)) {
                            $details.history = @(Parse-MeetingHistory -HistoryHtml $historyHtml)
                        }

                        Write-Host "OK ($($entry.date) $($details.timePeriod.resolvedTimeFrom)-$($details.timePeriod.resolvedTimeUntil) | Historie: $($details.history.Count))"
                    }
                    elseif ($entry.category -eq "ABSENCE") {
                        $details = Parse-AbsenceModal -ModalHtml $modalHtml
                        $entry.absenceDetails = $details
                        Set-ExactTimesFromAbsence -Entry $entry -Details $details
                        Write-Host "OK ($($details.dateStart) bis $($details.dateStop) | $($details.absenceReasonName))"
                    }
                    elseif ($entry.category -eq "SPECIALTIME") {
                        $details = Parse-SpecialtimeModal -ModalHtml $modalHtml -FallbackEntry $entry
                        $entry.specialtimeDetails = $details
                        Set-ExactTimesFromSpecialtime -Entry $entry -Details $details
                        Write-Host "OK ($($entry.date) $($details.timePeriod.resolvedTimeFrom)-$($details.timePeriod.resolvedTimeUntil) | $($details.specialtimeTypeName))"
                    }
                }
                else {
                    Write-Host "FEHLER (keine Antwort)" -ForegroundColor Red
                }

                Start-Sleep -Milliseconds 100
            }

            $allFetchedMonths.Add($parsedCalendar)

            # Gesamt-JSON speichern
            $jsonFilePath = Join-Path $desktopPath "DF_Kalender_$monthTag.json"
            $jsonOutput   = $parsedCalendar | ConvertTo-Json -Depth 12
            Save-TextToFile -FilePath $jsonFilePath -Content $jsonOutput -Label "`nGesamt-JSON $monthTag ($($parsedCalendar.users.Count) Mitarbeiter, $($allEntries.Count) Detail-Datensätze)"
        }
    }
    finally {
        if ($session -and $session.Client) {
            $session.Client.Dispose()
        }
    }

    # 5. Eigene Termine in lokales Outlook übertragen
    if ($allFetchedMonths.Count -gt 0) {
        Invoke-ConfigureAndSyncOutlook -Exports $allFetchedMonths -AvailableUsers $allAvailableUsers -Config $config -CredentialsPath $credentialsPath
    }
}

# Skript starten
Invoke-Main
Write-Host "`nFertig. Drücke eine beliebige Taste zum Beenden..."
[Console]::ReadKey($true) | Out-Null
