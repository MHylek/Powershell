function run2007{
echo "Running 2007 Script"
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
cls
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
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">600 West Fulton Street</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">Chicago, Illinois 60661-1259</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">D '+$ADTelephoneNumber+'</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><a href="mailto:"+$ADEmailAddress+">'+$ADEmailAddress+'</a></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><a href="Http://www.epsteinglobal.com">www.epsteinglobal.com</a></div><br>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><b>Mission</b></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><font color = "C75B12">Epstein</font> is a multi-disciplinary design and construction company focused on </div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">serving our clients, empowering our employee-owners, and enhancing our communities.</div>'))
$stream.WriteLine(("</body>"))
$stream.WriteLine(("</HTML>"))
$stream.close()
cls
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
cls
echo "Second Signature Complete!"
}


function setsignatures {
echo "Setting Default Signature..."
$wrd = new-object -com word.application 

$EmailOptions = $wrd.EmailOptions
$EmailSignature = $wrd.EmailOptions.EmailSignature
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
$EmailSignature.NewMessageSignature="epsteinsignature"
$EmailSignature.ReplyMessageSignature="epsteinReplysignature"
cls
echo "Default Signature Set!"
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
$wrd = new-object -com word.application -ErrorAction Stop
 
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

echo "Running 2007 Script"
### PROCESS ###
#STEP 1
createprimarysignature
#STEP 2
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


Invoke-Item "C:\Program files (X86)\Microsoft Office\Office12\outlook.exe"
}
function run2013{

echo "Running 2013 Script"
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
cls
echo "Folders Moved!"
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
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">600 West Fulton Street</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">Chicago, Illinois 60661-1259</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">D '+$ADTelephoneNumber+'</div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><a href="mailto:"+$ADEmailAddress+">'+$ADEmailAddress+'</a></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><a href="Http://www.epsteinglobal.com">www.epsteinglobal.com</a></div><br>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><b>Mission</b></div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial"><font color = "C75B12">Epstein</font> is a multi-disciplinary design and construction company focused on </div>'))
$stream.WriteLine(('<div style= "font-size:8pt; font-family:Arial">serving our clients, empowering our employee-owners, and enhancing our communities.</div>'))
$stream.WriteLine(("</body>"))
$stream.WriteLine(("</HTML>"))
$stream.close()
cls
echo "Primary signature Complete"
Start-Sleep -s 3
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
cls
echo "Secondary Signature Complete"
Start-Sleep -s 3
}
function setsignatures {
echo "Setting Default Signature..."
$wrd = new-object -com word.application 

$EmailOptions = $wrd.EmailOptions
$EmailSignature = $wrd.EmailOptions.EmailSignature
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
$EmailSignature.NewMessageSignature="epsteinsignature"
$EmailSignature.ReplyMessageSignature="epsteinReplysignature"
cls
echo " Signature set"
Start-Sleep -s 3

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
cls
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

### PROCESS ###
#STEP 1
createprimarysignature

#STEP 2
createSecondarySignature
#STEP 3
convertprimary
#Step 4
convertsecondary
#Step 5
copyfolders
#Step 6
setsignatures
Set-Location C:\
echo "Complete!"
Start-Sleep -s 5





}

run2013
