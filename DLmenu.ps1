# might need to install exchange online management as a pre-req... maybe check here?

#Connect-ExchangeOnline

#load Exchange functions from other script
. ".\ExchangeFunctions.ps1"

Add-Type -AssemblyName System.Windows.Forms

#Create menu window
$form = New-Object System.Windows.Forms.Form
$form.Text = "Add User to DLs"
$form.Width = 300
$form.Height = 300
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

$checkbox2 = New-Object System.Windows.Forms.CheckBox
$checkbox2.Location = New-Object System.Drawing.Point(50, 90)
$checkbox2.Size = New-Object System.Drawing.Size(200, 30)
$checkbox2.Text = "Basel"
$form.Controls.Add($checkbox2)

$checkbox3 = New-Object System.Windows.Forms.CheckBox
$checkbox3.Location = New-Object System.Drawing.Point(50, 120)
$checkbox3.Size = New-Object System.Drawing.Size(200, 30)
$checkbox3.Text = "Research Group"
$form.Controls.Add($checkbox3)

#click to submit
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(100, 160)
$button.Size = New-Object System.Drawing.Size(100, 30)
$button.Text = "Submit"
$form.Controls.Add($button)

$button.Add_Click({
    $username = $textBox.Text
    $selectedOptions = @()
    if ($checkbox1.Checked) { $selectedOptions += $checkbox1.Text }
    if ($checkbox2.Checked) { $selectedOptions += $checkbox2.Text }
    if ($checkbox3.Checked) { $selectedOptions += $checkbox3.Text }
    
    foreach ($groupName in $selectedOptions)
    {
        Add-DLUser $groupName $username
    }
    #[System.Windows.Forms.MessageBox]::Show("Selected options: " + ($selectedOptions -join ", "))
})

[void]$form.ShowDialog()