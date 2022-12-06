Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.speech

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Text-To-Speech'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.ShowIcon=$False
$form.StartPosition = 'CenterScreen'
$form.Topmost = $true

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,10)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$textBox.Text = "Hello world!"
$form.Controls.Add($textBox)

$btn1 = New-Object System.Windows.Forms.Button
$btn1.Location = New-Object System.Drawing.Point(10,35)
$btn1.Size = New-Object System.Drawing.Size(75,23)
$btn1.Text = 'Say'
$btn1.Add_Click(
{
    Add-Type -AssemblyName System.speech
    $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $speak.Speak($textBox.Text)
})

$btn2 = New-Object System.Windows.Forms.Button
$btn2.Location = New-Object System.Drawing.Point(85,35)
$btn2.Size = New-Object System.Drawing.Size(75,23)
$btn2.Text = 'Save'
$btn2.Add_Click(
{
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.forms")
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.FileName = $textBox.Text # Default file name
    $dlg.DefaultExt = ".txt" # Default file extension
    $dlg.Filter = "Waveform files|*.wav" # Filter files by extension

    # Show save file dialog box
    $result = $dlg.ShowDialog()

    # Process save file dialog box results
    if ($result) {
        # Save document
        $filename = $dlg.FileName
        Add-Type -AssemblyName System.speech

        $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
        $speak.SetOutputToWaveFile($filename)
        $speak.Speak($textBox.Text)
        $speak.Dispose()
        start $filename
    }    
})

$form.Controls.Add($btn1)
$form.Controls.Add($btn2)

$result = $form.ShowDialog()