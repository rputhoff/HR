Import-Module ActiveDirectory

function Scramble-String {
    param([string]$InputString)
    $chars = $InputString.ToCharArray()
    $rand = New-Object System.Random
    for ($i = $chars.Length - 1; $i -gt 0; $i--) {
        $j = $rand.Next(0, $i + 1)
        $temp = $chars[$i]
        $chars[$i] = $chars[$j]
        $chars[$j] = $temp
    }
    -join $chars
}

# Prompt for username
$user = Read-Host "Enter the username (sAMAccountName)"

# Prompt for password (since AD passwords cannot be retrieved)
$password = Read-Host "Enter the password for $user" -AsSecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
$plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)

# Scramble the password
$scrambled = Scramble-String -InputString $plainPassword

# Convert scrambled password to SecureString
$scrambledSecure = ConvertTo-SecureString -String $scrambled -AsPlainText -Force

# Set the scrambled password in Active Directory
Set-ADAccountPassword -Identity $user -NewPassword $scrambledSecure -Reset
