<#
    Find what exe files exist in c:\windows\system32 but not c:\windows\syswow64 and vice versa.
    Inspired because quser.exe does not exist in a 32 bit binary form

    @guyrleech 07/07/2020
#>

[int]$nativeCount = 0
[int]$32bitCount = 0

dir -path ([Environment]::GetFolderPath( [System.Environment+SpecialFolder]::System )) -Filter *.exe |ForEach-Object `
{
    $nativeCount++
    if( ! (Test-Path( ([System.IO.Path]::Combine( [Environment]::GetFolderPath( [System.Environment+SpecialFolder]::SystemX86 ) , $_.Name )))))
    {
        $32bitCount++
    }
}

"$nativeCount native exe files in $([Environment]::GetFolderPath( [System.Environment+SpecialFolder]::System )) of these $32bitCount 32 bit exes are not in $([Environment]::GetFolderPath( [System.Environment+SpecialFolder]::SystemX86 ))"

$32bitCount = $nativeCount = 0

[array]$32bitOnly = @( dir -path ([Environment]::GetFolderPath( [System.Environment+SpecialFolder]::SystemX86 )) -Filter *.exe |ForEach-Object `
{
    $32bitCount++
    if( ! (Test-Path( ([System.IO.Path]::Combine( [Environment]::GetFolderPath( [System.Environment+SpecialFolder]::System ) , $_.Name )))))
    {
        $_
    }
})

"$32bitCount 32 bit exe files in $([Environment]::GetFolderPath( [System.Environment+SpecialFolder]::SystemX86 )) of these $($32bitOnly.Count) 64 bit exes are not in $([Environment]::GetFolderPath( [System.Environment+SpecialFolder]::System ))"

$32bitOnly
