Set-Location "$env:USERPROFILE\send_applysys_email\venv\Scripts"
.\Activate.ps1
Set-Location "$env:USERPROFILE\send_applysys_email\app\"
python "main.py"
