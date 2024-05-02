Add-PSSnapin *EXC*
$Certificates = Get-ExchangeCertificate
if ($Certificates)
{
$Date = (Get-Date).AddDays(-30)
$Certificate = Get-ExchangeCertificate | Where-Object {$_.NotAfter -LT $Date -and $_.FriendlyName -NotLike "*MS-Organization*"}

$Output = @()
Foreach ($C in $Certificate)
{
    $Output += [PSCustomObject]@{
        FriendlyName = $C.FriendlyName
        Subject = $C.Subject
        Thumbprint = $C.Thumbprint
        NotAfter = $C.NotAfter.ToString("dd-MM-yyyy")
        Services = $C.Services       
}
    }
        }
Else
{
$Output += [PSCustomObject]@{
        Error = "Get-ExchangeCertificate returns null"
        Fix = "https://www.alitajran.com/get-exchangecertificate-blank-output/"
}
    }
$Output
