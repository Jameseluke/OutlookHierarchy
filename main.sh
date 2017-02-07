function getEmployeesByName{
  # $query contains the names of the employees to search for in format "LastName, FirstName"
  $queries = $args[0]
  # if $correct is true, will not check if query and result match
  $correctflag = $args[2]
  $logfile = $args[1]
  $failure = ""
  $temp = New-Object System.Collections.ArrayList
  $outlook = New-Object -ComObject Outlook.Application
  foreach($query in $queries){
    $correct = $correctflag
    if ($correct){
      $emp = $outlook.Session.GetGlobalAddressList().AddressEntries.Item($query)
    }
    while(!$correct){
      $emp = $outlook.Session.GetGlobalAddressList().AddressEntries.Item($query)
      $match = checkMatch $query $emp.Name
      if($match -eq -1){
        $nomatch = $query
        $query = Read-Host "Enter a new name to search for (Leave blank to skip)"
        if (!$query){
          $failure += " $nomatch`r`n"
          break
        }
      }
      $correct = if ($match -eq -1) {$FALSE} else {$TRUE}
    }
    if ($correct){
      $temp.Add($emp) | Out-Null
    }
  }
  if($failure){
    $errorString = "Names without matches:`r`n" + $failure
    $errorString | Out-File -encoding ASCII $logfile
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
  $log = [System.IO.Path]::GetTempFileName()
  $unproccessedEmp = getEmployeesByName $queries $log $correct
  $success = "Name, UID, Department, Role, Manager`r`n"
  $employees = @{}
  Clear-Host
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
  $out = $args[1]
  $success | Out-File -encoding ASCII $out
  Write-Host "Finished, Output saved to $out"
  Get-Content $log | Write-Host
  Get-Item $log | Remove-Item
}
