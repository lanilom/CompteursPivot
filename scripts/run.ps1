# ##########################################################################################################
#
# Migration_FiabilisationBDDAssurés_compteursPivot
#
# ##########################################################################################################
#
# OUTFILE
# - NUMASSU_CONTROLE: 
#    - NORA: [Code du produit (NORA)] + right([N° de compte],8)
#
# - NUMASSU_PIVOT
# - €_MNTISU
# - €_MNTSUP
# - €_NBPENG
#
# - €_CONTROLE
#   - NORA: [Valeur de rachat fin N-1 (ind.)]
#
# - €_MNTISU-€_CONTROLE : [€_MNTISU]-[€_CONTROLE]
# - €_MNTISU-€_CONTROLE > 1 : if abs([€_MNTISU-€_CONTROLE]) > 1 then "VRAI" else "" endif
# - UC_MNTISU
# - UC_MNTSUP
# - UC_NBPENG
#
# - UC_CONTROLE
#   - NORA: [Valorisation de la garantie UC au 31/12/AA-1]
#
# - UC_MNTISU-UC_CONTROLE : [UC_MNTISU]-[UC_CONTROLE]
# - UC_MNTISU-UC_CONTROLE > 1 : if abs([UC_MNTISU-UC_CONTROLE]) > 1 then "VRAI" else "" endif
#
# - AbsentDuPivot : IF [NUMASSU_PIVOT]="" THEN "VRAI" ELSE "" ENDIF
#
# - AbsentDuFichierDeControle : IF [NUMASSU_CONTROLE]="" THEN "VRAI" ELSE "" ENDIF
# 
# 
# 
#
# ##########################################################################################################
$processOPERATIONSOnly = $false




cls

cd $PSScriptRoot

. .\common\unzip.ps1
. .\common\loadHashTableFromCSVfile.ps1
. .\specific\getFileTemplateNameFromFilename.ps1
. .\specific\getApplicationFromFilename.ps1
#. .\specific\loadReferencePM.ps1
. .\specific\loadListOfAllreadyProcessedFiles.ps1
. .\specific\batch_counters.ps1
. .\specific\operations_counters.ps1
. .\specific\operations_counters2.ps1

$moduleDir = split-path -Path $pwd.ProviderPath -parent


$inDir = "$moduleDir\in"
$outDir = "$moduleDir\out"
$tempDir = "$moduleDir\temp"
$refDir = "$moduleDir\ref"
$arcDir = "$moduleDir\archives"


$COLSEP = ";"

# load list of allready processed files

$allreadyProcessedZipFiles = loadListOfAllreadyProcessedFiles "$outDir\report.csv" $COLSEP


# load transcodification table from CDPRDTNORA  to CDPRDT
# $hash_CDPRDTNORA2CDPRDT = loadHashTableFromCSVfile "$refDir\CDPRDTNORA2CDPRDT.csv" $COLSEP "CDPRDTNORA" "CDPRDT"



# remove temporary files
if ( Test-Path -path $tempDir ) { Remove-Item -path "$tempDir\*" }

#####

foreach ($zipFile in Get-Item -Path "$inDir\*.zip") {

    # skip if file has already been processed
    if ( $allreadyProcessedZipFiles.ContainsKey( $zipFile.Name ) -eq $true ) {
        
        $msg = "The file {0} has already been processed" -f $zipFile.Name
        Write-host $msg -f DarkYellow
    
    } else { # process zipfile otherwise

        write-host $zipFile.Name

        unzip2 $zipFile.FullName $tempDir

#####
    
        $hash_batchCounters = @{}
        $hash_operations_counters = @{}
        $hash_operations_counters2 = @{}
        $hash_referencePM = @{}

        # compute counters
        foreach ( $file in Get-Item -Path "$tempDir\*.csv"  ) {

            "START PROCESSING {0}" -f $file.Name | write-host 

            # get file template name

            $fileTemplateName = getFileTemplateNameFromFilename $file.Name

            # if $processOPERATIONSOnly -eq $true -> process only OPERATIONS files
            if ( ( $fileTemplateName -ne "OPERATIONS" ) -and ( $processOPERATIONSOnly -eq $true ) ) {
                continue;
            }


            $APPLICATION = getApplicationFromFilename $file.Name
            $NUMLOT = ""

            # reset batch counter for $fileTemplateName
            $hash_batchCounters[ $fileTemplateName ]+=0


            Import-Csv -Path $file.FullName -Delimiter ";" | % {

                # count nb of lines

                $hash_batchCounters[ $fileTemplateName ] += 1

                $hash_batchCounters[ "NBLIGN" ]+= 1

                if ( $fileTemplateName -eq "OPERATIONS" ) {

                    # if it is the first line of the OPERATIONS -> then get NUMLOT value and load PM references files

                    $operation = $_

                    if ( $hash_batchCounters[ "OPERATIONS" ] -eq 1 ) {

                        $NUMLOT = $operation.NUMLOT

                        # $hash_reference_PM is passed by reference
#                        loadReferencePM $APPLICATION $NUMLOT $hash_referencePM

                    }

                    # update batch counters
                    update_batch_counters $hash_batchCounters $operation

                    # update OPERATIONS counters
                    update_operations_counters $hash_operations_counters $operation


                    # #############################################
                    # update OPERATIONS counters at the NMCPT level
                    # #############################################
#                    update_operations_counters2 $hash_operations_counters2 $hash_CDPRDTNORA2CDPRDT $hash_referencePM $operation

                }

            }

        }
    }

    # add in hash_operations_counters2 accounts with reference PM and not in the OPERATIONS file

#    add_accounts_with_referencePM_and_not_in_OPERATIONS_file $hash_referencePM $hash_operations_counters2

    # ###########################################
    # print batch counters
    # ###########################################
    print_batch_counters  $hash_batchCounters     

        
    # ###########################################
    # print operations counters
    # ###########################################
#        print_operations_counters $hash_operations_counters


    # ###########################################
    # print operations counters2
    # ###########################################
#    print_operations_counters2 $hash_operations_counters2 




    # remove temporary files
    if ( Test-Path -path $tempDir ) {Remove-Item -path "$tempDir\*"}
####

    # add counters to reportfile

    # move zip file into $arcDir
    Move-Item $zipFile.FullName $arcDir
####

    write-host

}
