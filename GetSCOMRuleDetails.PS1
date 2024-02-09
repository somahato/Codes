$outputFile = “c:\temp\userRoles.txt”

$userRoles =  Get-SCOMUserRole

ForEach ($userRole in $userRoles){
“Role: ” + $userRole.DisplayName >> $outputFile
“Desc: ” + $userRole.Description >> $outputFile
“Members” >> $outputFile
“——-” >> $outputFile
ForEach ($usr in $userRole.Users){
  “`t” + $usr >> $outputFile
}
“Group Scope” >> $outputFile
“———–” >> $outputFile
ForEach ($class in $userRole.Scope.Objects){
  “`t” + (Get-SCOMClass -id $class).Name >> $outputFile
}

“==============================” >> $outputFile
“” >> $outputFile
“” >> $outputFile

}
