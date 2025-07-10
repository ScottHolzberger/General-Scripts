Add-Type -AssemblyName System.Windows.Forms
$msgBox = [Windows.Forms.MessageBox]
$msgBox2 = [Windows.Forms.MessageBox]
$buttonYes = [System.Windows.Forms.MessageboxDefaultButton]::Button1
$buttonNo = [System.Windows.Forms.MessageboxDefaultButton]::Button2
$buttonCancel = [System.Windows.Forms.MessageboxDefaultButton]::Button3

$msgBox = [System.Windows.Forms.MessageBox]::Show("I love you.  Do you love me?", 'I love you.', 'YesNo', 'Warning', 'Button2')


switch  ($msgBox) {
      'Yes' {
          [System.Windows.MessageBox]::Show("Hip Hip Horray", "Yeah")
          #Write-Host "YES"
      }

      'No' {
          [System.Windows.MessageBox]::Show("Oh Poopy Pants", "Oh ")
          #Write-Host "NO"
      }
  }