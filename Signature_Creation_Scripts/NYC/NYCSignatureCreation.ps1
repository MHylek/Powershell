function run2007{

[reflection.assembly]::LoadWithPartialName("Microsoft.Office.Interop.Word")
New-Item C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -type directory -force

 
function copyfolders {

echo "Moving Folders 60 Seconds left..."
Start-Sleep -s 10
Move-Item $FolderLocation\\epsteinsignature.htm C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force
Start-Sleep -s 10
Move-Item $FolderLocation\\epsteinReplysignature.htm C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force
echo "Moving Folders 40 Seconds left..." 
Start-Sleep -s 10
Move-Item $FolderLocation\\epsteinsignature.rtf C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force 
Start-Sleep -s 10
Move-Item $FolderLocation\\epsteinReplysignature.rtf C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force 
echo "Moving Folders 10 Seconds left..."
Start-Sleep -s 5
Move-Item $FolderLocation\\epsteinsignature.txt C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force 
Start-Sleep -s 5
Move-Item $FolderLocation\\epsteinReplysignature.txt C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force 
Start-Sleep -s 5
}
function createprimarysignature {
New-Item $FolderLocation -type directory -force
################################## Primary Signature ###########################################################

echo "Please Wait....Runnning..."
echo "Creating Signature..."
$stream = [System.IO.StreamWriter] "$FolderLocation\\epsteinsignature.htm"
$stream.WriteLine(("<HTML>"))
$stream.WriteLine(('<div style= "font-size:10pt; font-family:Arial"><strong>'+$ADDisplayName+'</strong></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">'+$ADTitle+'</div><br>'))
$stream.WriteLine(('<div style= "font-size:12pt; font-family:Arial"><strong><font color = "C75B12">EPSTEIN</font></strong></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><i>Building on Experience</i></div><br>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><b>Architecture | Interiors | Engineering | Construction</b></div><br>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">156 Ludlow Street, 4th Floor</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">New York, NY 10002</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">D '+$ADTelephoneNumber+'</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><a href="mailto:"+$ADEmailAddress+">'+$ADEmailAddress+'</a></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><a href="Http://www.epsteinglobal.com">www.epsteinglobal.com</a></div><br>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><b>Mission</b></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><font color = "C75B12">Epstein</font> is a multi-disciplinary design and construction company focused on </div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">serving our clients, empowering our employee-owners, and enhancing our communities.</div>'))
$stream.WriteLine(("</body>"))
$stream.WriteLine(("</HTML>"))
$stream.close()

echo "Primary signature Complete"
}
function createSecondarySignature {

echo "Please Wait....Runnning..."
echo "Creating Second Signature..."
$stream = [System.IO.StreamWriter] "$FolderLocation\\epsteinReplysignature.htm"
$stream.WriteLine(("<HTML>"))
$stream.WriteLine(('<div style= "font-size:10pt; font-family:Arial"><strong>'+$ADDisplayName+'</strong></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">'+$ADTitle+'</div><br>'))
$stream.WriteLine(('<div style= "font-size:12pt; font-family:Arial"><strong><font color = "C75B12">EPSTEIN</font></strong></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">D '+$ADTelephoneNumber+'</div>'))
$stream.WriteLine(("</body>"))
$stream.WriteLine(("</HTML>"))
$stream.close()
}
function setsignatures {
echo "Setting Default Signature..."
$wrd = new-object -com word.application 

$EmailOptions = $wrd.EmailOptions
$EmailSignature = $wrd.EmailOptions.EmailSignature
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
$EmailSignature.NewMessageSignature="epsteinsignature"
$EmailSignature.ReplyMessageSignature="epsteinReplysignature"
}
function convertprimary{
$wrd = new-object -com word.application 
 
# Make Word Visible 
$wrd.visible = $false 
 
# Open a document  
$doc = $wrd.documents.open('C:\scripts\epsteinsignature.htm') 

 #Save as rtf
$opt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatrtf
$name = 'C:\scripts\epsteinsignature.rtf'
$wrd.ActiveDocument.Saveas($name,$opt)

# Save as txt
$opt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatTextLineBreaks
$name = 'C:\scripts\epsteinsignature.txt'
$wrd.ActiveDocument.Saveas($name,$opt)

# Close word
$wrd.Quit()
echo "Waiting for Secondary Signature..."
Start-Sleep -s 10}
function convertsecondary{
$wrd = new-object -com word.application 
 
# Make Word Visible 
$wrd.visible = $false 
 
# Open a document  
$doc = $wrd.documents.open('C:\scripts\epsteinReplysignature.htm') 

 #Save as rtf
$opt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatrtf
$name = 'C:\scripts\epsteinReplysignature.rtf'
$wrd.ActiveDocument.Saveas($name,$opt)

# Save as txt
$opt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatTextLineBreaks
$name = 'C:\scripts\epsteinReplysignature.txt'
$wrd.ActiveDocument.Saveas($name,$opt)

# Close word
$wrd.Quit()
}


 
#Get Active Directory information for current user 
$UserName = $env:username 
$Filter = “(&(objectCategory=User)(samAccountName=$UserName))” 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.Filter = $Filter 
$ADUserPath = $Searcher.FindOne() 
$ADUser = $ADUserPath.GetDirectoryEntry() 
$ADDisplayName = $ADUser.DisplayName
$ADDepartment = $ADUser.department 
$ADEmailAddress = $ADUser.mail 
$ADTitle = $ADUser.title 
$ADTelePhoneNumber = $ADUser.TelephoneNumber 
$ADDisplayName
$ADEmailAddress
$ADTitle
$ADTelePhoneNumber
$ADDepartment

$FolderLocation = 'c:\scripts'  

cls
### PROCESS ###
#STEP 1
createprimarysignature
#STEP 2
cls 
createSecondarySignature
#STEP 3
convertprimary
#Step 4
convertsecondary
#Step 5
copyfolders
#Step 6
setsignatures
cls
Set-Location C:\
echo "Complete!"
Start-Sleep -s 5


if( 'C:\Program Files (x86)\Microsoft Office\Office15\Outlook.exe'){
invoke-item 'C:\Program Files (x86)\Microsoft Office\Office15\Outlook.exe'} else {Invoke-Item 'C:\Programfiles (X86)\Microsoft Office\Office12\outlook.exe'}

}
function run2013{
[reflection.assembly]::LoadWithPartialName("Microsoft.Office.Interop.Word")
New-Item C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -type directory -force

 
function copyfolders {

echo "Moving Folders 60 Seconds left..."
Start-Sleep -s 10
Move-Item $FolderLocation\\epsteinsignature.htm C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force
Start-Sleep -s 10
Move-Item $FolderLocation\\epsteinReplysignature.htm C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force
echo "Moving Folders 40 Seconds left..." 
Start-Sleep -s 10
Move-Item $FolderLocation\\epsteinsignature.rtf C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force 
Start-Sleep -s 10
Move-Item $FolderLocation\\epsteinReplysignature.rtf C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force 
echo "Moving Folders 10 Seconds left..."
Start-Sleep -s 5
Move-Item $FolderLocation\\epsteinsignature.txt C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force 
Start-Sleep -s 5
Move-Item $FolderLocation\\epsteinReplysignature.txt C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures -force 
Start-Sleep -s 5
}
function createprimarysignature {
New-Item $FolderLocation -type directory -force
################################## Primary Signature ###########################################################

echo "Please Wait....Runnning..."
echo "Creating Signature..."
$stream = [System.IO.StreamWriter] "$FolderLocation\\epsteinsignature.htm"
$stream.WriteLine(("<HTML>"))
$stream.WriteLine(('<div style= "font-size:10pt; font-family:Arial"><strong>'+$ADDisplayName+'</strong></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">'+$ADTitle+'</div><br>'))
$stream.WriteLine(('<div style= "font-size:12pt; font-family:Arial"><strong><font color = "C75B12">EPSTEIN</font></strong></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><i>Building on Experience</i></div><br>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><b>Architecture | Interiors | Engineering | Construction</b></div><br>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">156 Ludlow St, 4th Floor</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">New York, New York 10002</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">D '+$ADTelephoneNumber+'</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><a href="mailto:"+$ADEmailAddress+">'+$ADEmailAddress+'</a></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><a href="Http://www.epsteinglobal.com">www.epsteinglobal.com</a></div><br>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><b>Mission</b></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><font color = "C75B12">Epstein</font> is a multi-disciplinary design and construction company focused on </div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">serving our clients, empowering our employee-owners, and enhancing our communities.</div>'))
$stream.WriteLine(("</body>"))
$stream.WriteLine(("</HTML>"))
$stream.close()

echo "Primary signature Complete"
}
function createSecondarySignature {

echo "Please Wait....Runnning..."
echo "Creating Second Signature..."
$stream = [System.IO.StreamWriter] "$FolderLocation\\epsteinReplysignature.htm"
$stream.WriteLine(("<HTML>"))
$stream.WriteLine(('<div style= "font-size:10pt; font-family:Arial"><strong>'+$ADDisplayName+'</strong></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">'+$ADTitle+'</div><br>'))
$stream.WriteLine(('<div style= "font-size:12pt; font-family:Arial"><strong><font color = "C75B12">EPSTEIN</font></strong></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">D '+$ADTelephoneNumber+'</div>'))
$stream.WriteLine(("</body>"))
$stream.WriteLine(("</HTML>"))
$stream.close()
}
function setsignatures {
echo "Setting Default Signature..."
$wrd = new-object -com word.application 

$EmailOptions = $wrd.EmailOptions
$EmailSignature = $wrd.EmailOptions.EmailSignature
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
$EmailSignature.NewMessageSignature="epsteinsignature"
$EmailSignature.ReplyMessageSignature="epsteinReplysignature"
}
function convertprimary{
$wrd = new-object -com word.application 
 
# Make Word Visible 
$wrd.visible = $false 
 
# Open a document  
$doc = $wrd.documents.open('C:\scripts\epsteinsignature.htm') 

 #Save as rtf
$opt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatrtf
$name = 'C:\scripts\epsteinsignature.rtf'
$wrd.ActiveDocument.Saveas([ref]$name,[ref]$opt)

# Save as txt
$opt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatTextLineBreaks
$name = 'C:\scripts\epsteinsignature.txt'
$wrd.ActiveDocument.Saveas([ref]$name,[ref]$opt)

# Close word
$wrd.Quit()
echo "Waiting for Secondary Signature..."
Start-Sleep -s 10}
function convertsecondary{
$wrd = new-object -com word.application 
 
# Make Word Visible 
$wrd.visible = $false 
 
# Open a document  
$doc = $wrd.documents.open('C:\scripts\epsteinReplysignature.htm') 

 #Save as rtf
$opt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatrtf
$name = 'C:\scripts\epsteinReplysignature.rtf'
$wrd.ActiveDocument.Saveas([ref]$name,[ref]$opt)

# Save as txt
$opt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatTextLineBreaks
$name = 'C:\scripts\epsteinReplysignature.txt'
$wrd.ActiveDocument.Saveas([ref]$name,[ref]$opt)

# Close word
$wrd.Quit()
}


 
#Get Active Directory information for current user 
$UserName = $env:username 
$Filter = “(&(objectCategory=User)(samAccountName=$UserName))” 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.Filter = $Filter 
$ADUserPath = $Searcher.FindOne() 
$ADUser = $ADUserPath.GetDirectoryEntry() 
$ADDisplayName = $ADUser.DisplayName
$ADDepartment = $ADUser.department 
$ADEmailAddress = $ADUser.mail 
$ADTitle = $ADUser.title 
$ADTelePhoneNumber = $ADUser.TelephoneNumber 
$ADDisplayName
$ADEmailAddress
$ADTitle
$ADTelePhoneNumber
$ADDepartment

$FolderLocation = 'c:\scripts'  

cls
### PROCESS ###
#STEP 1
createprimarysignature
#STEP 2
cls 
createSecondarySignature
#STEP 3
convertprimary
#Step 4
convertsecondary
#Step 5
copyfolders
#Step 6
setsignatures
cls
Set-Location C:\
echo "Complete!"
Start-Sleep -s 5


if( 'C:\Program Files (x86)\Microsoft Office\Office15\Outlook.exe'){
invoke-item 'C:\Program Files (x86)\Microsoft Office\Office15\Outlook.exe'} else {Invoke-Item 'C:\Programfiles (X86)\Microsoft Office\Office12\outlook.exe'}

}




if($PSVersionTable.PSVersion.Major -eq 2){run2013}
elseif($PSVersionTable.PSVersion.Major -eq 4){echo "Running Failover" 
run2013}else{run2007

cls
Echo "Ran 2007 Script"}