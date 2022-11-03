$SearchBase = 'OU=,OU= ,OU=,DC=,DC='

$data = @(
    Get-AdUser -SearchBase $SearchBase -Filter * -Properties Name,UserPrincipalName,pager,lastLogon | 
    select Name,UserPrincipalName,pager,lastLogon
)


# Созадём объект Excel
$Excel = New-Object -ComObject Excel.Application

# Делаем его видимым
$Excel.Visible = $true
$WorkBook = $Excel.Workbooks.Add()

$ad_user = $WorkBook.Worksheets.Item(1)
$ad_user.Name = 'ad_user'

$ad_user.Cells.Item(1,1) = 'Name'
$ad_user.Cells.Item(1,2) = 'UserPrincipalName'
$ad_user.Cells.Item(1,3) = 'pager'
$ad_user.Cells.Item(1,4) = 'lastLogon'

#перебор и запись
$count = 1
foreach($user in $data){
    $count += 1;

    $ad_user.Cells.Item($count,1) = $user.Name
    $ad_user.Cells.Item($count,2) = $user.UserPrincipalName
    $ad_user.Cells.Item($count,3) = $user.pager
    $ad_user.Cells.Item($count,4) = $user.lastLogon
}


#сохраняем закрываем
$WorkBook.SaveAs('C:\Scripts\Report.xlsx')
$Excel.Quit()
