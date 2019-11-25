<# 

.SYNOPSIS
	Applicable Staff Photo Uploader for Office 365.
	
.DESCRIPTION
	This script prompts you for the staff photo to be used with Office 365 which will sync across the `n
    Office 365 tenant and show up on Skype and Exchange using Set-UserPhoto Cmdlets.
	
.EXAMPLE
	Just run this script with no additional parameters.
 
.NOTES

 Copyright © 2017 Applicable Ltd
 
 Author: Adnan Khan <ads.khan@applicable.com>
 
 ChangeLog v1.0 - Ads Khan - 2017.02.17 - Initial Script


#>


<# 
We'll use the Microsoft .NET framework to create a system object for the open file dialogue form.
This will be shown so the user can interactively select the photo to be uploaded.
We'll later call this function below.
#>

Function Grab-File($startingdirectory)

{
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

    # Create the dialogue box object
    $dialogue = New-Object System.Windows.Forms.OpenFileDialog

    #The starting directory which will be displayed when the form pops up.
    $dialogue.InitialDirectory = $startingdirectory

    # We're going to ensure only JPG files show up using the filter below for security and to avoid breaking anything.
    $dialogue.filter = "JPG (*.jpg)| *.jpg"

    #Present the dialogue box to the user to select the photo
    $dialogue.ShowDialog() | Out-Null
    $dialogue.FileName
    

}




Write-Host -ForegroundColor Yellow "You're about to connect to the Office 365 tenant. You will be prompted for credentials.`nPlease just use your applicable email and password for this"
Pause

#Create variable to store user credentials to connect to the tenant.
$creds = get-credential

#Create a new session with URI to point to online powershell virtual directory.
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/?proxymethod=rps" -Credential $creds -Authentication Basic -AllowRedirection

#Import the session above so we can run commands from it.
Import-PSSession $session
cls

$user = Read-Host "Please enter the email address of the user you're going to assign a photo for."
Write-Host -ForegroundColor Yellow "You will now be prompted to select the photo you're going to use for this user.`n`nPlease ensure it is no larger than 30kb in file size and cropped to square`nformat with max resolution of 648x648."
Pause

#Call the function above to browse for user photo to be user. Passed string paramter is c:\ as the initial directory.
#Photo variable will be used to hold the content to read bytes from.
$photo = Grab-File -startingdirectory "c:\"



#Set the user's photo to what was chosen in the function 
#We will not be using the -confirm flag to $false as we want the user to be able to confirm
#they are sure they want to perform this operation.
Set-UserPhoto -Identity "$user"  -verbose -PictureData ([System.IO.File]::ReadAllBytes($photo))
Write-Host -Foreground Yellow "If all went well and you received no warnings this should have worked.`nLook up the user in Skype to check if it's taken effect.`n`nYou might need to wait a few minutes"

#End
Write-Host -ForegroundColor Yellow "End Of Script"