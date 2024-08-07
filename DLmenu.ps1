# might need to install exchange online management as a pre-req... maybe check here?

#load Exchange functions from other script
. ".\ExchangeFunctions.ps1"

Add-Type -AssemblyName System.Windows.Forms

$checkBoxList = @()

#Create menu window
$form = New-Object System.Windows.Forms.Form
$form.Text = "Add User to DLs"
$form.Width = 300
$form.Height = 350
$form.StartPosition = "CenterScreen"

#Create username input field
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,5)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please enter the users email in the space below:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,30)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

#create buttons to select DLs
$checkbox1 = New-Object System.Windows.Forms.CheckBox
$checkbox1.Location = New-Object System.Drawing.Point(50, 60)
$checkbox1.Size = New-Object System.Drawing.Size(200, 30)
$checkbox1.Text = "All"
$form.Controls.Add($checkbox1)
$checkBoxList += $checkbox1

$checkbox2 = New-Object System.Windows.Forms.CheckBox
$checkbox2.Location = New-Object System.Drawing.Point(50, 90)
$checkbox2.Size = New-Object System.Drawing.Size(200, 30)
$checkbox2.Text = "Basel"
$form.Controls.Add($checkbox2)
$checkBoxList += $checkbox2

$checkbox3 = New-Object System.Windows.Forms.CheckBox
$checkbox3.Location = New-Object System.Drawing.Point(50, 120)
$checkbox3.Size = New-Object System.Drawing.Size(200, 30)
$checkbox3.Text = "Research Group"
$form.Controls.Add($checkbox3)
$checkBoxList += $checkbox3

$checkbox4 = New-Object System.Windows.Forms.CheckBox
$checkbox4.Location = New-Object System.Drawing.Point(50, 150)
$checkbox4.Size = New-Object System.Drawing.Size(200, 30)
$checkbox4.Text = "San Diego"
$form.Controls.Add($checkbox4)
$checkBoxList += $checkbox4

$checkbox5 = New-Object System.Windows.Forms.CheckBox
$checkbox5.Location = New-Object System.Drawing.Point(50, 180)
$checkbox5.Size = New-Object System.Drawing.Size(200, 30)
$checkbox5.Text = "Stamford"
$form.Controls.Add($checkbox5)
$checkBoxList += $checkbox5

$checkbox6 = New-Object System.Windows.Forms.CheckBox
$checkbox6.Location = New-Object System.Drawing.Point(50, 210)
$checkbox6.Size = New-Object System.Drawing.Size(200, 30)
$checkbox6.Text = "Watertown"
$form.Controls.Add($checkbox6)
$checkBoxList += $checkbox6

#click to submit
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(100, 240)
$button.Size = New-Object System.Drawing.Size(100, 30)
$button.Text = "Submit"
$form.Controls.Add($button)

$button.Add_Click({
    $username = $textBox.Text
    foreach ( $checkboxName in $checkBoxList)
    {
        if ($checkboxName.Checked) {
            Add-DLUser $checkboxName.Text $username
        } else {
            Write-Host "$checkboxName is not checked"
        }
    }
})

[void]$form.ShowDialog()