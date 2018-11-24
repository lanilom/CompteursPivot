cd $PSScriptRoot

$files = @()


# retrieve input files
foreach ($file in Get-Item -Path ".\in\*.csv") {

    $files += $file.FullName

}


# error if more than one csv file
if ( $files.count -gt 1 ) {

    write-host "ERROR: more than one csv files are in  the .\in directory." -f DarkRed

    exit
}


# process single csv file

write-host "Processing file $files" -f Green

$line = 0
$header = ""
$yyMMdd_HHmmss = (Get-Date -Format yyMMdd-HHmmss)

Get-content $files[0] | % {

    $line++

    # process header
    if ($line -eq 1) {

        $_ > ".\out\NORA_1_ALL_$yyMMdd_HHmmss.csv"
        $_ > ".\out\NORA_2_ALL_$yyMMdd_HHmmss.csv"
        $_ > ".\out\NORA_3_ALL_$yyMMdd_HHmmss.csv"

    # process lines 2.. to end of file
    } else {

        $CDPRDT = $_.subString(0,3)
        $lot = "4"

        # lot 1
        if ( $CDPRDT -match "093|763|768" ) { $lot = "1" } 

        # lot 2
        if ( $CDPRDT -match "442|754|829|847" ) { $lot = "2" }

        # lot 3
        if ( $CDPRDT -match "710|813|817|818" ) { $lot = "3" }

        $outFilename = ".\out\NORA_{0}_ALL_{1}.csv" -f $lot, $yyMMdd_HHmmss

        $_ >> $outFilename

    }

}
