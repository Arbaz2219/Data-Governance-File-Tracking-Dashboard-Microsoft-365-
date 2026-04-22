# LDP LOGISTICS - DESKTOP INTELLIGENCE AGENT
# Run this script on any computer to track browser history for free.

$ServerUrl = "https://apildpm365.ldplogistics.com/api/security/report-activity" # Updated for production backend
$DeviceName = $env:COMPUTERNAME
$UserName = $env:USERNAME

Write-Host "🚀 LDP Desktop Intelligence Agent Started..." -ForegroundColor Cyan
Write-Host "Target: Tracking Edge/Chrome History"

while($true) {
    try {
        # 1. Paths to Browser History (Edge is typically identical to Chrome)
        $EdgePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\History"
        $ChromePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\History"
        $Paths = @($EdgePath, $ChromePath)

        foreach ($Path in $Paths) {
            if (Test-Path $Path) {
                # 2. Copy file to bypass 'In Use' lock
                $TempPath = "$env:TEMP\browser_hist_temp"
                Copy-Item -Path $Path -Destination $TempPath -Force

                # 3. Extract URLs using Regex (Zero-Dependency Method)
                # We look for strings that look like common URLs
                $RawData = Get-Content -Path $TempPath -Raw -Encoding Byte
                $StringData = [System.Text.Encoding]::UTF8.GetString($RawData)
                
                # Match common URLs (basic regex for dashboard purposes)
                $Matches = [regex]::Matches($StringData, 'https?://[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,}[^\s\x00-\x1F]*')
                
                # Get uniquely recent ones (last 5 to keep pings light)
                $UniqueUrls = $Matches.Value | Select-Object -Unique | Select-Object -Last 5

                foreach ($Url in $UniqueUrls) {
                    # 4. Detect Search Terms
                    $SearchTerm = ""
                    if ($Url -match "q=([^&]+)") { $SearchTerm = $Matches[0].Groups[1].Value }
                    if ($Url -match "query=([^&]+)") { $SearchTerm = $Matches[0].Groups[1].Value }

                    # 5. POST to Dashboard
                    $Payload = @{
                        deviceName = $DeviceName
                        accountName = $UserName
                        remoteUrl = $Url
                        searchTerm = $SearchTerm
                        timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    } | ConvertTo-Json

                    Invoke-RestMethod -Uri $ServerUrl -Method Post -Body $Payload -ContentType "application/json" -ErrorAction SilentlyContinue
                }
            }
        }
    } catch {
        # Fail silently to avoid interrupting the user
    }
    
    Write-Host "." -NoNewline
    Start-Sleep -Seconds 30 # Check every 30 seconds
}
