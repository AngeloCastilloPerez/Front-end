# =============================================================
#                        USER TWEAKY BITS
# =============================================================


# 1: Replace Name with the theme name you want for the SharPoint Theme colour
# Do not delete the double-quotes (" ")
$SPThemeName = "Theme ENAPU5"

# 2:  Paste the PowerShell code from the Theme Designer 
# after the equal sign below.
$SPThemePalette =@{
"themePrimary" = "#ffffff";
"themeLighterAlt" = "#767676";
"themeLighter" = "#a6a6a6";
"themeLight" = "#c8c8c8";
"themeTertiary" = "#d0d0d0";
"themeSecondary" = "#dadada";
"themeDarkAlt" = "#eaeaea";
"themeDark" = "#f4f4f4";
"themeDarker" = "#f8f8f8";
"neutralLighterAlt" = "#172332";
"neutralLighter" = "#162331";
"neutralLight" = "#16212f";
"neutralQuaternaryAlt" = "#141f2c";
"neutralQuaternary" = "#131e2a";
"neutralTertiaryAlt" = "#121c28";
"neutralTertiary" = "#fdf4eb";
"neutralSecondary" = "#fef6ee";
"neutralSecondaryAlt" = "#fef6ee";
"neutralPrimaryAlt" = "#fef8f1";
"neutralPrimary" = "#fcefe1";
"neutralDark" = "#fefbf8";
"black" = "#fffdfb";
"white" = "#182534";
}
# 3: Go to line 63 change the "domain" to your Office 365 sharepoint domain

# 4: Click File --> Save to save the changes to the script.

# 5: Click on the line below and press [F8]
Set-ExecutionPolicy RemoteSigned -Force

# 6: Either Press the play button |> in PowerShell ISE toolbar above
# or Press F5.
# =============================================================
     
 
# =============================================================
#                         THE CODE
# =============================================================
 
    cls
    if(Get-module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell") {
        Write-Host "SharePoint Online module already installed" -ForegroundColor Green
    }
    else {
        Write-Host "Installing latest version of SharePoint Online module" -ForegroundColor Yellow
        Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force
    }
    
    #  You need to supply an Office 365 account with SharePoint Admin rights
    Connect-SPOService -Url https://enapu1-admin.sharepoint.com/

    Add-SPOTheme -Identity $SPThemeName -Palette $SPThemePalette -IsInverted $false

    Write-Host "Your Fluent UI Theme '"$SPThemeName "' is now added to SharePoint.`n Go to a web site and Change the Look." -ForegroundColor Green

    Set-ExecutionPolicy Default -Force