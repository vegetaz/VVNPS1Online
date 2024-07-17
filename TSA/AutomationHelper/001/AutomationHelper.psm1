function Get-Greeting {
    param (
        [string]$Name
    )
    return "Hello, $Name!"
}

Export-ModuleMember -Function Get-Greeting
