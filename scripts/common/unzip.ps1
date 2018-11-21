# unzip
# ex: unzip "C:\a.zip" "C:\a"
# does not work on remote volume

Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip {
    param([string]$zipfile, [string]$outpath)

    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

# unzip2
# Doesn't work with windows 2012/2016 'core' edition
# If one of the files or directories already exists at the destination location, 
# it pops up a dialogue asking what to do (ignore, overwrite) which defeats the purpose. 
# Does anyone know how to force it to silently overwrite? –

function unzip2 {

    param([string]$zipfile, [string]$outpath)

    $shell = New-Object -ComObject shell.application

    $zip = $shell.NameSpace( $zipfile )

    foreach ($item in $zip.items()) {
        $shell.Namespace( $outpath ).CopyHere($item)
    }

}