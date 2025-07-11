# Pull hash information 
Get-Process -Id (Get-Process -Name "notepad").Id | Get-FileHash

