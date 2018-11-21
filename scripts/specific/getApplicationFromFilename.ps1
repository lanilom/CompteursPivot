
Function getApplicationFromFilename {

    param ( [string]$filename )

    $application = ""


    $fileNameSegments = $filename.Split("_")


    if ( $fileNameSegments[0] -eq "CNP2" ) {

        $application = "NORA"

    }


    if ( $fileNameSegments[0] -eq "CNP1" ) {

        $application = "LOIC"

    }

    $application

}