function getEmployeeByName{
  $query = $args[0]
  $outlook = New-Object -ComObject Outlook.Application
  return $outlook.Session.GetGlobalAddressList().AddressEntries.Item($query)
}

function checkMatch{
  $string1 = $args[0]
  $string2 = $args[1]
  $confirmation = Read-Host "$string1 == $string2 [y/n]"
  while($confirmation -ne "y")
  {
      if ($confirmation -eq 'n') {return -1}
      $confirmation = Read-Host "Please choose one of the available options [y/n]"
  }
  return 0
}

function main{
  param(
        [switch] $correct)
  $queries = Get-Content $args[0]
  $success = "Name, UID, Department, Role, Manager`r`n"
  $failure = "Names without matches:`r`n"
  $employees = @{}
  foreach($query in $queries){
    $employee = getEmployeeByName $query
    if(!$correct){
      $match = checkMatch $query $employee.Name
    }
    if(!$correct -and ($match -eq -1)){
      $failure += " $query`r`n"
    }
    else {
      $full = $employee.getExchangeUser()
      while($full.Alias -and !$employees.ContainsKey($full.Alias)){
        $temp = @{}
        $alias = $full.Alias
        $temp.Add("Name", "`"" + $full.Name +"`"")
        $temp.Add("Department", "`"" + $full.Department +"`"")
        $temp.Add("Role", "`"" + $full.JobTitle +"`"")
        $full = $full.Manager
        $temp.Add("Manager", $full.Alias)
        $employees.Add($alias, $temp)
      }
    }
  }
  foreach ($h in $employees.Keys) {
    $success += $employees.$h.Name + ","
    $success += $h  + ","
    $success += $employees.$h.Department  + ","
    $success += $employees.$h.Role  + ","
    $success += $employees.$h.Manager + "`n"
  }
  $success | Out-File -encoding ASCII hierarchy.txt
}
