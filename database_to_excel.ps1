# https://www.sharepointdiary.com/2016/06/upload-files-to-sharepoint-online-using-powershell.html
#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
  
#Variables for Processing
$WebUrl = "https://crescent.sharepoint.com/Sites/Sales/"
$LibraryName ="Documents"
$SourceFile="C:\SitesToCreate.csv"
$AdminName ="Salaudeen@crescent.com"
$AdminPassword ="password goes here"
  
#Setup Credentials to connect
$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,(ConvertTo-SecureString $AdminPassword -AsPlainText -Force))
  
#Set up the context
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl) 
$Context.Credentials = $Credentials
 
#Get the Library
$Library =  $Context.Web.Lists.GetByTitle($LibraryName)
 
#Get the file from disk
$FileStream = ([System.IO.FileInfo] (Get-Item $SourceFile)).OpenRead()
#Get File Name from source file path
$SourceFileName = Split-path $SourceFile -leaf
   
#sharepoint online upload file powershell
$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
$FileCreationInfo.Overwrite = $true
$FileCreationInfo.ContentStream = $FileStream
$FileCreationInfo.URL = $SourceFileName
$FileUploaded = $Library.RootFolder.Files.Add($FileCreationInfo)
  
#powershell upload single file to sharepoint online
$Context.Load($FileUploaded) 
$Context.ExecuteQuery() 
 
#Close file stream
$FileStream.Close()
  
write-host "File has been uploaded!"