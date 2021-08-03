$teams = Get-Clipboard -Format Text
$pass = 'fill://'
$short = $teams + $pass

invoke-item $teams