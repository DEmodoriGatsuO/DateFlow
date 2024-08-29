# .NETのWindowsフォームアセンブリをロード
Add-Type -AssemblyName System.Windows.Forms

# フォームと日付ピッカーの作成
$form = New-Object Windows.Forms.Form
$form.Text = "日付を選択してください"
$form.Width = 300
$form.Height = 200

$dateTimePicker = New-Object Windows.Forms.DateTimePicker
$dateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Long
$dateTimePicker.Width = 250
$dateTimePicker.Value = Get-Date

$okButton = New-Object Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Width = 100
$okButton.Top = 100
$okButton.Left = 100
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

$form.Controls.Add($dateTimePicker)
$form.Controls.Add($okButton)
$form.AcceptButton = $okButton

# フォームを表示し、ユーザーが選択した日付を取得
$result = $form.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedDate = $dateTimePicker.Value.ToString("yyyy年MM月dd日")
    Write-Output "選択された日付: $selectedDate"
}
