function doesfieldLabelBelongToCSVFileHeader {

    param (

        [string]$fieldLabel,
        [string]$fieldDelimiter,
        [string]$CSVfilefullName
        

    )
#    write-host $fieldLabel

    $return = $false

    # get file header
    $header = Get-Content -Path $CSVfilefullName -TotalCount 1

    $headerSegments = $header.split($fieldDelimiter)

    for ($i=0 ; $i -lt $headerSegments.Count ; $i++) {

#         write-host $headerSegments[$i]

        if ( $headerSegments[$i] -eq $fieldLabel ) {

            $return = $true
            break

        }

    }

    if ( $return -eq $false ) {

        'ERROR: field "{0}" is not present in file {1}.' -f $fieldLabel, $CSVfilefullName | write-host -f DarkRed
        'Header is "{0}".' -f $header | write-host -f DarkRed
    }


    $return

}