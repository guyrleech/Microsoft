#requires -version 4
<#
    Put contents of file into clipboard - designed for use by explorer's send to menu, hence we don't use named parameters

    @guyrleech 2018

    Modification History:

    16/07/2020  @guyrleech  Added support for image formats
#>

$command = Get-Command -Name Out-Null
$addedType = $null

$(ForEach( $file in $args )
{
    switch -regex ( ( [system.io.path]::GetExtension( $file ) -replace '^\.' ) )
    {
        'gif|jpg|jpeg|png|jfif|bmp|tif' `
        { 
            if( ! $addedType )
            {
                $addedType = Add-Type -AssemblyName System.Windows.Forms -PassThru
            }
            if( $bitmap = [System.Drawing.Image]::FromFile( $file ) )
            {
                [Windows.Forms.Clipboard]::SetImage( $bitmap )
            }
        }
        default `
        {
            $command = Get-Command -Name Set-Clipboard
            Get-Content -Path $file
        }
    }
}) | . $command
