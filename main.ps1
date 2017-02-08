function getEmployeesByName{
  # $query contains the names of the employees to search for in format "LastName, FirstName"
  $queries = $args[0]
  # if $correctflag is true, will not check if query and result match
  $correctflag = $args[2]
  # $logfile will be written to in the event of no match being found
  $logfile = $args[1]
  # $failure intially empty so that it can be tested in an if statement
  $failure = ""
  # $temp is an arraylist of all employees that were successfully matched
  $temp = New-Object System.Collections.ArrayList
  # If outlook is not open, then script cannot access global address list
  $outlook = New-Object -ComObject Outlook.Application
  if(!$outlook){
    Write-Host "ERROR: Please ensure Outlook is open and connected to the server"
    exit
  }
  foreach($query in $queries){
    # create temp correct flag so that it can be changed for each query
    $correct = $correctflag
    if ($correct){
      # Assumes the first employee found is correct, does no checking with user
      $emp = $outlook.Session.GetGlobalAddressList().AddressEntries.Item($query)
    }
    while(!$correct){
      $emp = $outlook.Session.GetGlobalAddressList().AddressEntries.Item($query)
      $match = checkMatch $query $emp.Name
      if(!$match){
        # keep name in case user does not wish to search further
        $nomatch = $query
        $query = Read-Host "Enter a new name to search for (Leave blank to skip)"
        # if a new query is not submitted, skip this employee
        if (!$query){
          # This will cause an error message to be printed at the end of script
          $failure += " $nomatch`r`n"
          break
        }
      }
      # Check to see if loop should be broken or not
      $correct = if (!$match) {$FALSE} else {$TRUE}
    }
    # if a match was found, add employee to output arraylist
    if ($correct){
      $temp.Add($emp) | Out-Null
    }
  }
  # If there was one or more names without match, write their names to log file
  if($failure){
    $errorString = "Names without matches:`r`n" + $failure
    $errorString | Out-File -encoding ASCII $logfile
  }
  return $temp
}

# Return 1 if user thinks the two strings match, else 0
function checkMatch{
  $string1 = $args[0]
  $string2 = $args[1]
  $confirmation = Read-Host "$string1 == $string2 [y/n]"
  while($confirmation -ne "y")
  {
      if ($confirmation -eq 'n') {return 0}
      $confirmation = Read-Host "Please choose one of the available options [y/n]"
  }
  return 1
}

function main{
  param(
        # correct switch causes script to ignore input validation
        [switch] $correct)
  # List of all names to search from
  $queries = Get-Content $args[0]
  $log = [System.IO.Path]::GetTempFileName()
  # list of all employees which were matched to input names
  $unproccessedEmp = getEmployeesByName $queries $log $correct
  $success = "Name, UID, Department, Role, Manager`r`n"
  $employees = @{}
  Clear-Host
  Write-Progress -Activity "Working..."
  foreach($emp in $unproccessedEmp){
    $full = $emp.getExchangeUser()
    # Create an object containing all details about the employee
    # Object is referenced by alias within the employees object
    # Stops when employee has no manager or manager has already been found before
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
  # format for CSV output
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
  # Print errors to screen and delete log file
  Get-Content $log | Write-Host
  Get-Item $log | Remove-Item
}
