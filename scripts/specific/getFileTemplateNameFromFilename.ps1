
Function getFileTemplateNameFromFilename {

    param ( [string]$filename )

    $fileTemplateName = ""

    $fileNameSegments = $filename.Split("_")

    for ( $i = 1; $i -lt ($fileNameSegments.Count -2) ; $i++ ) {
            
        if ( $i -eq 1 ) {

            $fileTemplateName = $fileNameSegments[ $i ]
            
        } else {

            $fileTemplateName = "{0}_{1}" -f $fileTemplateName, $fileNameSegments[ $i ]

        }
    }


    $fileTemplateName

}