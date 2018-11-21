Function loadHashTableFromCSVfile {

    param ( $fullFillename,
        $colSep,
        $keyColName,
        $valueColName 

    )

    $hash = @{}

    Import-csv -Path $fullFillename -Delimiter $colSep | % {

        Try {

            $hash.Add( $_.$keyColName, $_.$ValueColName )
    
        } Catch {

            $msg = "Error while Loading key/value ({0}/{1}) from file {2} into hash table." -f $keyColName, $valueColName, $fullFillename
            Write-host $msg -f Red

        }

    }

    $hash
}