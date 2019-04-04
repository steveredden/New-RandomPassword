function New-RandomPassword {
    <#
    .SYNOPSIS
    Generates a new random password

    .DESCRIPTION
    Generates a new random password using ASCII characters 32 (space) through 126 (tilde)

    .PARAMETER PasswordLength
    The length of the new password
    
    .PARAMETER MinUpperCase
    The minimum number of uppercase characters (65:90) that must be present in the new password
    To exclude uppercase characters in the new password input -1

    .PARAMETER MinLowerCase
    The minimum number of lowercase characters (97:122) that must be present in the new password
    To exclude lowercase characters in the new password input -1

    .PARAMETER MinDigit
    The minimum number of digit characters (48:57) that must be present in the new password
    To exclude digit characters in the new password input -1

    .PARAMETER MinSpecial
    The minimum number of special characters (32:47, 58:64, 91:96, 123:126) that must be present in the new password
    To exclude special characters in the new password input -1

    .PARAMETER Forbidden
    String of chracters that must not be present in the new password

    .PARAMETER PreventRepeatingChars
    Include to ensure no characters are repated next to one another

    .PARAMETER TimeoutSec
    Number of seconds to attempt to generate a new password
    Depending on your restrictions you may hit this

    .EXAMPLE
    New-RandomPassword
    New-RandomPassword -Forbidden " |``~"
    New-RandomPassword -PasswordLength 60 -MinUpperCase 5 -MinLowerCase 5 -MinDigit 2 -MinSpecial 5 -Forbidden " " -PreventRepeatingChars -TimeoutSec 30

    .OUTPUTS
    String

    #>
    [CmdletBinding()]
	Param(
	[Parameter(Mandatory = $false)]
        [int]$PasswordLength = 30,

        [Parameter(Mandatory = $false)]
        [int]$MinUpperCase = 2,

        [Parameter(Mandatory = $false)]
        [int]$MinLowerCase= 2,

        [Parameter(Mandatory = $false)]
        [int]$MinDigit = 1,

        [Parameter(Mandatory = $false)]
        [int]$MinSpecial = 2,

        [Parameter(Mandatory = $false)]
        [string]$Forbidden = " ",

        [Parameter(Mandatory = $false)]
        [switch]$PreventRepeatingChars,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutSec = 15
	)

    [bool]$PreventRepeats = ($PreventRepeatingChars.IsPresent)

    If ( $MinUpperCase+$MinLowerCase+$MinDigit+$MinSpecial -gt $PasswordLength ) { return $null }

    [char[]]$SpecialArray=([int][char]' '..[int][char]'/')  #32:47
    [char[]]$DigitArray=([int][char]'0'..[int][char]'9')    #48:57
    $SpecialArray += ([int][char]':'..[int][char]'@')       #58:64
    [char[]]$UpperArray=([int][char]'A'..[int][char]'Z')    #65:90
    $SpecialArray += ([int][char]'['..[int][char]'`')       #91:96
    [char[]]$LowerArray=([int][char]'a'..[int][char]'z')    #97:122
    $SpecialArray += ([int][char]'{'..[int][char]'~')       #123:126

    $CharacterPool = New-Object System.Collections.ArrayList
    If ( $MinUpperCase -ne -1 ) { Foreach ( $Char in $UpperArray ) { $CharacterPool.Add($Char) | Out-Null } }
    If ( $MinLowerCase -ne -1 ) { Foreach ( $Char in $LowerArray ) { $CharacterPool.Add($Char) | Out-Null } }
    If ( $MinDigit -ne -1 ) { Foreach ( $Char in $DigitArray ) { $CharacterPool.Add($Char) | Out-Null } }
    If ( $MinSpecial -ne -1 ) { Foreach ( $Char in $SpecialArray ) { $CharacterPool.Add($Char) | Out-Null } }
    Foreach ( $Char in [char[]]$Forbidden ) { $CharacterPool.Remove($Char) | Out-Null }

    If ( $CharacterPool.Count -lt 2 ) { return $null }

    $Valid = $false
    $Timeout = (Get-Date).AddSeconds($TimeoutSec)
	While ( ($Valid -eq $false) -and ((Get-Date) -lt $Timeout) ) {

        [string]$NewPassword = "" 
        For ( $i = 0; $i -lt $PasswordLength; $i++ ) {
            $NewPassword += $CharacterPool[(Get-Random -Minimum 0 -Maximum ($CharacterPool.Count-1))]
        }

        If ( ($PreventRepeats -eq $false) -or ($NewPassword -notmatch "(.)\1") ) {
            If ( [Regex]::Matches($NewPassword, "[0-9]").Count -ge $MinDigit) {
                If ( [Regex]::Matches($NewPassword, "[A-Z]").Count -ge $MinUpperCase) {
		    If ( [Regex]::Matches($NewPassword, "[a-z]").Count -ge $MinLowerCase) {
                        If ( [Regex]::Matches($NewPassword, "[$(-join $SpecialArray)]").Count -ge $MinSpecial) {
                            $Valid = $true
                        }
                    }
                }
            }
        }
    }
    If ( (Get-Date) -ge $Timeout ) { return $null }
    return $NewPassword
}
