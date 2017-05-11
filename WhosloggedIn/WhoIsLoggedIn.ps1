$computer = read-host "Computer name"
# Shows current users logged in
$current = quser /server:$computer

$current