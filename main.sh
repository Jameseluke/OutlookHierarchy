function getEmployeesByName{
  # $query contains the names of the employees to search for in format "LastName, FirstName"
  $queries = $args[0]
  # if $correct is true, will not check if query and result match
  $correct = $args[1]
  $temp = New-Object System.Collections.ArrayList
  $outlook = New-Object -ComObject Outlook.Application
  foreach($query in $queries){
    $emp = $outlook.Session.GetGlobalAddressList().AddressEntries.Item($query)
    if(!$correct){
      $match = checkMatch $query $emp.Name
    }
    if(!$correct -and ($match -eq -1)){
      $failure += " $query`r`n"
    }
    else {
      $temp.Add($emp) | Out-Null
    }
  }
  return $temp
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
  $unproccessedEmp = getEmployeesByName $queries $correct
  $success = "Name, UID, Department, Role, Manager`r`n"
  $failure = "Names without matches:`r`n"
  $employees = @{}
  cls
  Write-Progress -Activity "Working..."
  foreach($emp in $unproccessedEmp){
    $full = $emp.getExchangeUser()
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
  foreach ($h in $employees.Keys) {
    $success += $employees.$h.Name + ","
    $success += $h  + ","
    $success += $employees.$h.Department  + ","
    $success += $employees.$h.Role  + ","
    $success += $employees.$h.Manager + "`n"
  }
  $success | Out-File -encoding ASCII $args[1]
  echo "Finished, Output saved to $args[1]"
}
