<# 
Modified for Powershell 5.1 by Ben Cleverdon 
Originally Written By Eric Holzhueter with no implied guarantee it will work

This does not check for prerequisites, specifically the following note from MS

    Important Note: For systems running Windows 7 SP1 or Windows Server 2008 R2 SP1, you’ll need WMF 4.0 and .NET Framework 4.5 or higher installed to run WMF 5.0.

It also does not cleanup the file when done. Some might want this.

WMF 5.1 download page
   https://www.microsoft.com/en-us/download/details.aspx?id=54616

WMF Blog posts 
    https://blogs.msdn.microsoft.com/powershell/2017/01/19/windows-management-framework-wmf-5-1-released/

OS version table https://msdn.microsoft.com/en-us/library/windows/desktop/ms724832%28v=vs.85%29.aspx

#>

#Get the operating environment
$Arch = (Get-WmiObject -Class Win32_Processor).addresswidth
$OS = (Get-WmiObject Win32_OperatingSystem).version.split('.')[0]
$OSMinor = (Get-WmiObject Win32_OperatingSystem).version.split('.')[1]
$isZip = $false

#64 bit OS check
if($Arch -eq 64){
    if($OS -eq 6 -and $OSMinor -eq 1){$url = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win7AndW2K8R2-KB3191566-x64.zip"
                                      $file = "Win7AndW2K8R2-KB3191566-x64.zip"
                                      $isZip = $true}
    if($OS -eq 6 -and $OSMinor -eq 2){$url = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/W2K12-KB3191565-x64.msu"
                                      $file = "W2K12-KB3191565-x64.msu"}
    if($OS -eq 6 -and $OSMinor -eq 3){$url = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win8.1AndW2K12R2-KB3191564-x64.msu"
                                      $File = "Win8.1AndW2K12R2-KB3191564-x64.msu"}
}
elseif($Arch -eq 32){
    if($OS -eq 6 -and $OSMinor -eq 1){$url = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win7-KB3191566-x86.zip"
                                      $file = "Win7-KB3191566-x86.zip"
                                      $isZip = $true}
    if($OS -eq 8 -and $OSMinor -eq 3){$url = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win8.1-KB3191564-x86.msu"
                                      $file = "Win8.1-KB3191564-x86.msu"}
}

if($url -eq $null -or $file -eq $null){Write-Error "Failed to find valid OS"}
Else{
    $output = Join-Path $env:TEMP $file

    Write-Output "Downloading: $file"
    Write-Output "From: $url"
    Write-Output "To: $output"

    $wc = New-Object System.Net.Webclient
    #May be required in some places with authenticated proxy
    #$wc.UseDefaultCredentials = $true
    #$wc.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
    $wc.DownloadFile($url, $output)
    
    #Unzip Windows 7/Server 2008 using PoSh 3 syntax
    If ($output -match ".zip"){
        
        Add-Type -assembly "system.io.compression.filesystem"
        $destination = Split-Path "$output"
        $destination = "$destination\output"
        [io.compression.zipfile]::ExtractToDirectory($output, $destination)
        
        #Run Install-WMF5.1
        $command= "$destination\Install-WMF5.1.ps1 -AcceptEULA"
        Invoke-Expression $command
        }
    Else{    
        if(test-path $output){ Write-Output "Installing $file"
                               Start-Process wusa -ArgumentList ($output, '/quiet', '/norestart', "/log:$env:HOMEDRIVE\Temp\Wusa.log") -Wait
        else{Write-Warning "File $((Join-Path $output $file)) does not exist"}
      }
}
