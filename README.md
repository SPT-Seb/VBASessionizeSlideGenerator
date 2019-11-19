# VBASessionizeSlideGenerator
VBA project used to populate a PowerPoint PPTX file based on session content extracted at Excel format from Sessionize (see https://sessionize.com/ ) portal. 
Build for aOS Community usage (see https://aos.community/) to generate pictures for each speakers during events
By the way, this is a good way to understand how to use VBA to provision PowerPoint content ;)

To test it, download sources and open a Powershell prompt from local folder.
Run command  .\Format-SessionizeExcelToPowerPoint.ps1 -excelFileFullPath '.\sessionize - data sample.xlsx' -powerPointFileFullPath '.\TEST Annonce Speaker.pptx'

Requires : 
- PowerShell (for command confort, you can run it directly from VBA)
- VBA
- Office 365 Pro Plus (should run on classical Office version but not tested)
