# -- Constants -- #
$olFormatUnspecified = 0
$olFormatPlain = 1
$olFormatHTML = 2
$olFormatRichText = 3

$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = "Receiver"
$Mail.CC = "Carbon copy"
$Mail.Subject = "Subject"
$Mail.Bodyformat = $olFormatHTML
$Mail.HTMLBody = @"
	Body (formatted as HTML)
"@

$Mail.Send()
Out-Null