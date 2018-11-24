Function createDirectory {
    param (
        [string]$dir=""
    )



    if ( ( Test-path -Path $dir ) -eq $false ) { New-Item -ItemType directory -Path $dir }

    return $dir
}

Function copyFile {
    param (
        [string]$sourcefile="",
        [string]$targetDir=""
    )

    $sourceFileSegments = $sourceFile.Split($DIRSEP)

    $targetFile = "{0}\{1}" -f $targetDir, $sourceFileSegments[ ($sourceFileSegments.Count - 1) ]

    if ( (Test-path -Path $targetFile) -eq $false ) { Copy-Item -Path $sourcefile -Destination $targetFile }


}


# #####################################################
#  TESTING
# #####################################################

$DIRSEP="\"

clear

cd $PSScriptRoot

. ..\..\..\scripts\specific\operations_counters2.ps1
. ..\..\..\scripts\common\load_XMLCONFIG.ps1
. ..\..\..\scripts\common\doesfieldLabelBelongToCSVFileHeader.ps1

# create test directories

$rootDir= ( createDirectory "c:\temp\loadReferencePM_NORA" )
$refDir = ( createDirectory "c:\temp\loadReferencePM_NORA\ref" )
$tempDir = ( createDirectory "c:\temp\loadReferencePM_NORA\temp" )
$outDir = ( createDirectory "c:\temp\loadReferencePM_NORA\out" )
$configDir = ( createDirectory "c:\temp\loadReferencePM_NORA\config" )
$logDir = ( createDirectory "c:\temp\loadReferencePM_NORA\log" )




# set test files

copyFile .\ref\NORA_1_ALL_20181022.csv $refDir
copyFile .\temp\CNP2_OPERATIONS_12112018_V0.1.csv $tempDir
copyFile .\config\config.xml $configDir


# set variables

[hashtable]$hash = @{}

$XMLCONFIG = load_XMLCONFIG $configDir $logDir

write-host "Mapping des champs:"
$xmlconfig.SelectNodes('//config/fields/field') | % {

    $msg = '{0,-5} {1,-3 } {2,-25} {3,-35}' -f $_.application, $_.produit, $_.code, $_.label 

    $msg

}


# APPLICATION: NORA | LOIC
$APPLICATION = "NORA"

# NUMLOT: 1 | 2 | 3 
$NUMLOT = "1" 


# test load reference PM
Measure-Command {
    loadReferencePM $APPLICATION $NUMLOT $hash
}


# test load operations PM
foreach ( $file in Get-Item -Path "$tempDir\*_OPERATIONS_*.csv"  ) {

    Measure-Command {
        loadOperationsPM $file.FullName $hash
    }

}

   

# test print_operations_counters2
$zipfile = $file

Measure-Command {
    print_operations_counters2 $APPLICATION $NUMLOT $hash
}

    