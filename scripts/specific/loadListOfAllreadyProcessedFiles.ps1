Function loadListOfAllreadyProcessedFiles {

    param ( $fullFilename, 
            $colsep
    )

    $hash = @{}

    if ( Test-path -Path $fullFilename ) {

        Get-Content -Path $fullFilename  | % {

            $cols = $_.ToString().Split( $colsep )

            $hash.Add( $cols[0] ,"")

        }
    }

    $hash

}