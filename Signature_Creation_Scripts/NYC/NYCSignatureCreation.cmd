taskkill /im outlook.exe /f
mkdir C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Signatures


powershell -NoProfile -executionPolicy Bypass -command "& '\\chifile01\apps\Windows Common\Scripts\Signature_Creation_Scripts\NYC\NYCSignatureCreation.ps1'"


