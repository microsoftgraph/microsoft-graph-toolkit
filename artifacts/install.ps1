$root = Split-Path -parent $PSCommandPath

function Get-Packages {
  param (
    [Parameter(
      Mandatory=$true,
      Position=0,
      ValueFromRemainingArguments=$true
    )]
    [string[]]
    $packageNames
  )

  [System.Collections.ArrayList]$output = @();
  foreach ($name in $packageNames) {
    $package = Get-ChildItem $root | Where Name -Match "microsoft-$name-[0-9]"
    $_ = $output.Add("$root\$package")
  }

  Write-Output $output
}

$packages = Get-Packages mgt-element mgt mgt-react

npm i $packages
