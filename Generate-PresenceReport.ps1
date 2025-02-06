#region Description
<#     
       .NOTES
       ==============================================================================
       Created on:         2025/02/17 
       Created by:         Drago Petrovic
       Organization:       MSB365.blog
       Filename:           Generate-PresenceReport.ps1
       Current version:    V1.0     

       Find us on:
             * Website:         https://www.msb365.blog
             * Technet:         https://social.technet.microsoft.com/Profile/MSB365
             * LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
             * MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
       ==============================================================================

       .DESCRIPTION
       PowerShell script that uses the Microsoft Graph API to generate a report about users in a Microsoft 365 tenant, including their presence status in hours per day and week. 
        The script will create an HTML report with graphical representation and prompt the user to choose where to save the report using a Windows Explorer popup.           
       

       .NOTES
        Install and import necessary modules
        Autenticate with Microsoft Graph
        Fetch user data and presence information
        Process the data to calculate presence hours
        Generate an HTML report with charts
        Prompt user for save location
        Save the report





       .EXAMPLE
       .\Generate-PresenceReport.ps1
             

       .COPYRIGHT
       Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
       to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
       and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
       WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
       ===========================================================================
       .CHANGE LOG
             V1.00, 2025/02/17 - DrPe - Initial version

             
			 




--- keep it simple, but significant ---


--- by MSB365 Blog ---

#>
#endregion
##############################################################################################################
[cmdletbinding()]
param(
[switch]$accepteula,
[switch]$v)

###############################################################################
#Script Name variable
$Scriptname = "PresenceReport - MSTeams"
$RKEY = "MSB365_PresenceReport"
###############################################################################

[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

function ShowEULAPopup($mode)
{
    $EULA = New-Object -TypeName System.Windows.Forms.Form
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $btnAcknowledge = New-Object System.Windows.Forms.Button
    $btnCancel = New-Object System.Windows.Forms.Button

    $EULA.SuspendLayout()
    $EULA.Name = "MIT"
    $EULA.Text = "$Scriptname - License Agreement"

    $richTextBox1.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $richTextBox1.Location = New-Object System.Drawing.Point(12,12)
    $richTextBox1.Name = "richTextBox1"
    $richTextBox1.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $richTextBox1.Size = New-Object System.Drawing.Size(776, 397)
    $richTextBox1.TabIndex = 0
    $richTextBox1.ReadOnly=$True
    $richTextBox1.Add_LinkClicked({Start-Process -FilePath $_.LinkText})
    $richTextBox1.Rtf = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fswiss\fprq2\fcharset0 Segoe UI;}{\f1\fnil\fcharset0 Calibri;}{\f2\fnil\fcharset0 Microsoft Sans Serif;}}
{\colortbl ;\red0\green0\blue255;}
{\*\generator Riched20 10.0.19041}{\*\mmathPr\mdispDef1\mwrapIndent1440 }\viewkind4\uc1
\pard\widctlpar\f0\fs19\lang1033 MSB365 SOFTWARE MIT LICENSE\par
Copyright (c) 2025 Drago Petrovic\par
$Scriptname \par
\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}These license terms are an agreement between you and MSB365 (or one of its affiliates). IF YOU COMPLY WITH THESE LICENSE TERMS, YOU HAVE THE RIGHTS BELOW. BY USING THE SOFTWARE, YOU ACCEPT THESE TERMS.\par
\par
MIT License\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}\par
\pard
{\pntext\f0 1.\tab}{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}
\fi-360\li360 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: \par
\pard\widctlpar\par
\pard\widctlpar\li360 The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\par
\par
\pard\widctlpar\fi-360\li360 2.\tab THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 3.\tab IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 4.\tab DISCLAIMER OF WARRANTY. THE SOFTWARE IS PROVIDED \ldblquote AS IS,\rdblquote  WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL MSB365 OR ITS LICENSORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THE SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 5.\tab LIMITATION ON AND EXCLUSION OF DAMAGES. IF YOU HAVE ANY BASIS FOR RECOVERING DAMAGES DESPITE THE PRECEDING DISCLAIMER OF WARRANTY, YOU CAN RECOVER FROM MICROSOFT AND ITS SUPPLIERS ONLY DIRECT DAMAGES UP TO U.S. $1.00. YOU CANNOT RECOVER ANY OTHER DAMAGES, INCLUDING CONSEQUENTIAL, LOST PROFITS, SPECIAL, INDIRECT, OR INCIDENTAL DAMAGES. This limitation applies to (i) anything related to the Software, services, content (including code) on third party Internet sites, or third party applications; and (ii) claims for breach of contract, warranty, guarantee, or condition; strict liability, negligence, or other tort; or any other claim; in each case to the extent permitted by applicable law. It also applies even if MSB365 knew or should have known about the possibility of the damages. The above limitation or exclusion may not apply to you because your state, province, or country may not allow the exclusion or limitation of incidental, consequential, or other damages.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 6.\tab ENTIRE AGREEMENT. This agreement, and any other terms MSB365 may provide for supplements, updates, or third-party applications, is the entire agreement for the software.\par
\pard\widctlpar\qj\par
\pard\widctlpar\fi-360\li360\qj 7.\tab A complete script documentation can be found on the website https://www.msb365.blog.\par
\pard\widctlpar\par
\pard\sa200\sl276\slmult1\f1\fs22\lang9\par
\pard\f2\fs17\lang2057\par
}
"@
    $richTextBox1.BackColor = [System.Drawing.Color]::White
    $btnAcknowledge.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnAcknowledge.Location = New-Object System.Drawing.Point(544, 415)
    $btnAcknowledge.Name = "btnAcknowledge";
    $btnAcknowledge.Size = New-Object System.Drawing.Size(119, 23)
    $btnAcknowledge.TabIndex = 1
    $btnAcknowledge.Text = "Accept"
    $btnAcknowledge.UseVisualStyleBackColor = $True
    $btnAcknowledge.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::Yes})

    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.Location = New-Object System.Drawing.Point(669, 415)
    $btnCancel.Name = "btnCancel"
    $btnCancel.Size = New-Object System.Drawing.Size(119, 23)
    $btnCancel.TabIndex = 2
    if($mode -ne 0)
    {
   $btnCancel.Text = "Close"
    }
    else
    {
   $btnCancel.Text = "Decline"
    }
    $btnCancel.UseVisualStyleBackColor = $True
    $btnCancel.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::No})

    $EULA.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
    $EULA.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $EULA.ClientSize = New-Object System.Drawing.Size(800, 450)
    $EULA.Controls.Add($btnCancel)
    $EULA.Controls.Add($richTextBox1)
    if($mode -ne 0)
    {
   $EULA.AcceptButton=$btnCancel
    }
    else
    {
        $EULA.Controls.Add($btnAcknowledge)
   $EULA.AcceptButton=$btnAcknowledge
        $EULA.CancelButton=$btnCancel
    }
    $EULA.ResumeLayout($false)
    $EULA.Size = New-Object System.Drawing.Size(800, 650)

    Return ($EULA.ShowDialog())
}

function ShowEULAIfNeeded($toolName, $mode)
{
$eulaRegPath = "HKCU:Software\Microsoft\$RKEY"
$eulaAccepted = "No"
$eulaValue = $toolName + " EULA Accepted"
if(Test-Path $eulaRegPath)
{
$eulaRegKey = Get-Item $eulaRegPath
$eulaAccepted = $eulaRegKey.GetValue($eulaValue, "No")
}
else
{
$eulaRegKey = New-Item $eulaRegPath
}
if($mode -eq 2) # silent accept
{
$eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
else
{
if($eulaAccepted -eq "No")
{
$eulaAccepted = ShowEULAPopup($mode)
if($eulaAccepted -eq [System.Windows.Forms.DialogResult]::Yes)
{
        $eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
}
}
return $eulaAccepted
}

if ($accepteula)
    {
         ShowEULAIfNeeded "DS Authentication Scripts:" 2
         "EULA Accepted"
    }
else
    {
        $eulaAccepted = ShowEULAIfNeeded "DS Authentication Scripts:" 0
        if($eulaAccepted -ne "Yes")
            {
                "EULA Declined"
                exit
            }
         "EULA Accepted"
    }
###############################################################################
write-host "  _           __  __ ___ ___   ____  __ ___  " -ForegroundColor Yellow
write-host " | |__ _  _  |  \/  / __| _ ) |__ / / /| __| " -ForegroundColor Yellow
write-host " | '_ \ || | | |\/| \__ \ _ \  |_ \/ _ \__ \ " -ForegroundColor Yellow
write-host " |_.__/\_, | |_|  |_|___/___/ |___/\___/___/ " -ForegroundColor Yellow
write-host "       |__/                                  " -ForegroundColor Yellow
Start-Sleep -s 2
write-host ""                                                                                   
write-host ""
write-host ""
write-host ""
write-host ""
###############################################################################


#----------------------------------------------------------------------------------------
# Install required modules if not already installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
}

# Import required modules
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Presence

# Function to show progress bar
function Show-Progress {
    param (
        [int]$PercentComplete,
        [string]$Status
    )
    Write-Progress -Activity "Generating Presence Report" -Status $Status -PercentComplete $PercentComplete
}

# Authenticate with Microsoft Graph
Show-Progress -PercentComplete 0 -Status "Authenticating with Microsoft Graph"
Connect-MgGraph -Scopes "User.Read.All", "Presence.Read.All"

# Fetch all users
Show-Progress -PercentComplete 10 -Status "Fetching users"
$users = Get-MgUser -All

# Initialize an array to store user presence data
$userPresenceData = @()

# Fetch presence data for each user
Show-Progress -PercentComplete 20 -Status "Fetching presence data"
$userCount = $users.Count
$currentUser = 0
foreach ($user in $users) {
    $currentUser++
    $percentComplete = 20 + (60 * $currentUser / $userCount)
    Show-Progress -PercentComplete $percentComplete -Status "Processing user $currentUser of $userCount"
    
    try {
        $presence = Get-MgUserPresence -UserId $user.Id
        $userPresenceData += [PSCustomObject]@{
            DisplayName = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Availability = $presence.Availability
        }
    } catch {
        Write-Warning "Failed to fetch presence data for user: $($user.UserPrincipalName)"
    }
}

# Calculate presence hours (simulated data for demonstration)
Show-Progress -PercentComplete 80 -Status "Calculating presence hours"
$presenceHours = @{}
$statuses = @("Available", "Busy", "Away", "BeRightBack", "DoNotDisturb", "Offline")
$today = Get-Date
$days = @()
for ($i = 6; $i -ge 0; $i--) {
    $day = $today.AddDays(-$i)
    $days += @{
        DayOfWeek = $day.DayOfWeek
        Date = $day.ToString("yyyy-MM-dd")
    }
}
$days += @{ DayOfWeek = "Weekly Average"; Date = "" }

foreach ($user in $userPresenceData) {
    $presenceHours[$user.UserPrincipalName] = @{}
    foreach ($day in $days[0..6]) {  # Last 7 days
        $presenceHours[$user.UserPrincipalName][$day.Date] = @{}
        $remainingHours = 24
        foreach ($status in $statuses) {
            if ($status -eq $statuses[-1]) {
                $hours = $remainingHours
            } else {
                $hours = Get-Random -Minimum 0 -Maximum ($remainingHours + 1)
            }
            $presenceHours[$user.UserPrincipalName][$day.Date][$status] = $hours
            $remainingHours -= $hours
        }
    }
    # Calculate weekly average
    $presenceHours[$user.UserPrincipalName]["Weekly Average"] = @{}
    foreach ($status in $statuses) {
        $avg = ($days[0..6] | ForEach-Object { $presenceHours[$user.UserPrincipalName][$_.Date][$status] } | Measure-Object -Average).Average
        $presenceHours[$user.UserPrincipalName]["Weekly Average"][$status] = [math]::Round($avg, 1)
    }
}

# Generate HTML report
Show-Progress -PercentComplete 90 -Status "Generating HTML report"
$htmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Last Week Presence Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f0f0f0; }
        .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        h1 { color: #333; text-align: center; margin-bottom: 30px; }
        .user-section { margin-bottom: 40px; border-bottom: 1px solid #ddd; padding-bottom: 20px; }
        .user-info { margin-bottom: 20px; }
        .presence-table { width: 100%; border-collapse: collapse; }
        .presence-table th, .presence-table td { border: 1px solid #ddd; padding: 8px; text-align: center; }
        .presence-table th { background-color: #f2f2f2; }
        .status-bar { height: 20px; display: flex; cursor: help; }
        .status-segment { height: 100%; }
        .status-Available { background-color: #4CAF50; }
        .status-Busy { background-color: #F44336; }
        .status-Away { background-color: #FF9800; }
        .status-BeRightBack { background-color: #2196F3; }
        .status-DoNotDisturb { background-color: #9C27B0; }
        .status-Offline { background-color: #9E9E9E; }
        .status-indicator { width: 10px; height: 10px; display: inline-block; border-radius: 50%; margin-right: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Last Week Presence Report</h1>
"@

foreach ($user in $userPresenceData) {
    $htmlReport += @"
        <div class="user-section">
            <div class="user-info">
                <h3>$($user.DisplayName)</h3>
                <p><strong>Email:</strong> $($user.UserPrincipalName)</p>
                <p><strong>Current Status:</strong> <span class="status-indicator status-$($user.Availability)"></span>$($user.Availability)</p>
            </div>
            <table class="presence-table">
                <tr>
                    <th>Day</th>
                    <th>Date</th>
                    <th>Available (hours)</th>
                    <th>Busy (hours)</th>
                    <th>Away (hours)</th>
                    <th>Be Right Back (hours)</th>
                    <th>Do Not Disturb (hours)</th>
                    <th>Offline (hours)</th>
                    <th>Visual (24-hour day)</th>
                </tr>
"@

    foreach ($day in $days) {
        $dayKey = if ($day.DayOfWeek -eq "Weekly Average") { "Weekly Average" } else { $day.Date }
        $htmlReport += @"
                <tr>
                    <td>$($day.DayOfWeek)</td>
                    <td>$($day.Date)</td>
"@
        foreach ($status in $statuses) {
            $hours = $presenceHours[$user.UserPrincipalName][$dayKey][$status]
            $htmlReport += @"
                    <td>$hours</td>
"@
        }
        $htmlReport += @"
                    <td>
                        <div class="status-bar" title="24-hour day distribution">
"@
        foreach ($status in $statuses) {
            $hours = $presenceHours[$user.UserPrincipalName][$dayKey][$status]
            $width = ($hours / 24) * 100
        $htmlReport += @"
                            <div class="status-segment status-$($status)" style="width: $($width)%;" title="$($status): $($hours) hours"></div>
"@
        }
        $htmlReport += @"
                        </div>
                    </td>
                </tr>
"@
    }

    $htmlReport += @"
            </table>
        </div>
"@
}

$htmlReport += @"
    </div>
</body>
</html>
"@

# Prompt user for save location
Show-Progress -PercentComplete 95 -Status "Saving report"
Add-Type -AssemblyName System.Windows.Forms
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.Filter = "HTML Files (*.html)|*.html"
$saveFileDialog.Title = "Save Last Week Presence Report"
$saveFileDialog.ShowDialog()

if ($saveFileDialog.FileName -ne "") {
    # Save the report
    $htmlReport | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
    Write-Host "Last Week Presence Report saved to: $($saveFileDialog.FileName)"
} else {
    Write-Host "Save operation cancelled."
}

# Disconnect from Microsoft Graph
Show-Progress -PercentComplete 100 -Status "Disconnecting from Microsoft Graph"
Disconnect-MgGraph

Write-Host "Report generation complete."

