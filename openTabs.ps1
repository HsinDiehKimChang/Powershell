
###Written by Hsin-Dieh (Kim) Chang
###2019/06/14
###Version 1



function menu
{
      cls
      Write-Host "=======Select page you want to open========="      
      Write-Host "1: Press '1' for open OUTLOOK."
      Write-Host "2: Press '2' for open ServiceNow."
      Write-Host "3: Press '3' for open your work note(excel)"
      Write-Host "q: Press 'q' to quit."
      
 }

  do
{
      menu
      $input = Read-Host "Please make a selection....."
      switch ($input)
      {
            '1' {
            Start-Process OUTLOOK
            }'2' {
            Start-Process "chrome.exe" https://service-now.com
            }'3' {
            $objExcel = New-Object -ComObject Excel.Application
            $objExcel.Visible = $true
            $FilePath='\\officescchome20.office.adroot.bmogc.net\scc20userdata$\hchan18\Desktop\Note\workNote.xlsx'
            $workbook = $objExcel.Workbooks.Open($FilePath)
            #Start-Process EXCEL $FilePath
            }
      }
     

}
until ($input -eq 'q') 
