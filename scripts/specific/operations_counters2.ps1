Function IsAbsoluteGapGreaterThan1 {

    param (
        [double]$value1=0,
        [double]$value2=0

    )

    $return = ""

    [double]$gap = ($value1 - $value2)

    if ( [math]::Abs( $gap ) -gt 1 ) {

        $return = "VRAI"

    } 

    $return

}

Function loadReferencePM_LOIC_813 {

    param (
        $fileFullname,
        [hashtable]$hash

    )

    write-host "   Start loadReferencePM_LOIC_813"

    Measure-Command {


        $MNTPM_TOTAL = 0

        Import-Csv -Path $file.FullName -Delimiter ";" -Encoding Default  | % {

            $APPLICATION = "LOIC"
            $CDPRDT = ($_."Ident compte LOIC").substring(0, 3)
            $CDCLLCTVT = ($_."Ident collectivité").substring(6, 5)
            $CDCNTRT = $_."Num contrat"
            $NMCPT = $_."Ident compte LOIC"
            $MNTPM =  [double]($_."Epargne au 31/12/n-1".Replace(",","."))

            $MNTPM_TOTAL += $MNTPM

            $key = "{0};{1};{2};{3};{4}" -f  $APPLICATION, $CDPRDT, $CDCLLCTVT, $CDCNTRT, $NMCPT

            if ( $hash.ContainsKey( $key )  )  {

                $hash[ $key ] += $MNTPM

            } else {

                $hash.Add( $key, $MNTPM)
            }
        }

        Write-host "TOTAL PM $fileFullname : $MNTPM_TOTAL" 

    }

    write-host "   END loadReferencePM_LOIC_813"

}


Function loadReferencePM_LOIC_818 {

    param (
        $fileFullname,
        [hashtable]$hash

    )

    write-host "   START loadReferencePM_LOIC_818"

    Measure-Command {

        $MNTPM_TOTAL = 0

        Import-Csv -Path $file.FullName -Delimiter ";" -Encoding Default  | % {

            $APPLICATION = "LOIC"
            $CDPRDT = ($_."Ident compte LOIC").substring(0, 3)
            $CDCLLCTVT = ($_."Ident collectivité").substring(6, 5)
            $CDCNTRT = $_."Num contrat"
            $NMCPT = $_."Ident compte LOIC"
            $MNTPM =  [double]($_."Mnt compte épargne revalorisé".Replace(",",".")) 
            $MNTPM += [double]($_."Mnt compte épargne revalorisé salar".Replace(",",".")) 

            $MNTPM_TOTAL += $MNTPM

            $key = "{0};{1};{2};{3};{4}" -f  $APPLICATION, $CDPRDT, $CDCLLCTVT, $CDCNTRT, $NMCPT

            if ( $hash.ContainsKey( $key )  )  {

                $hash[ $key ] += $MNTPM

            } else {

                $hash.Add( $key, $MNTPM)
            }
        }
    

        Write-host "TOTAL PM $fileFullname : $MNTPM_TOTAL" 

    }

    write-host "   END loadReferencePM_LOIC_818"

}


Function loadReferencePM_NORA {

    param (
        $fileFullname,
        [hashtable]$hash

    )

    $TOTALPM_EURO = 0
    $TOTALPM_UC = 0

    Import-Csv -Path $file.FullName -Delimiter ";" -Encoding Default  | % {

        # compute NUMASSU

        $CDPRDT_NORA = $_."Code du produit (NORA)"
        $NMCPT = $_."N° de compte".Substring(3,8)
        $NUMASSU = "{0}{1}" -f $CDPRDT_NORA, $NMCPT

        # get EURO and UC values

        # EURO = [Valeur de rachat fin N-1 (ind.)]
        $EURO_CONTROLE = [double]($_."Valeur de rachat fin N-1 (ind.)".Replace(",","."))

        # UC = [Valorisation de la garantie UC au 31/12/AA-1]
        $UC_CONTROLE = [double]($_."Valorisation de la garantie UC au 31/12/AA-1".Replace(",","."))


        # compute total PM
        $TOTALPM_EURO += $EURO_CONTROLE
        $TOTALPM_UC += $UC_CONTROLE


        # search if NUMASSU was already inserted in the index


        if ( $hash.ContainsKey( $NUMASSU ) ) { # update NUMASSU € and UC related values

            if ( $EURO_CONTROLE -gt 0 ) {
                $hash[$NUMASSU]["€_CONTROLE"] += $EURO_CONTROLE
                $hash[$NUMASSU]["€_MNTISU-€_CONTROLE"] = (0 - $hash[$NUMASSU]["€_CONTROLE"])
                $hash[$NUMASSU]["€_MNTISU-€_CONTROLE > 1"] = IsAbsoluteGapGreaterThan1 0 $hash[$NUMASSU]["€_CONTROLE"]
            }

            if ( $UC_CONTROLE -gt 0 ) {
                $hash[$NUMASSU]["UC_CONTROLE"] += $UC_CONTROLE
                $hash[$NUMASSU]["UC_MNTISU-UC_CONTROLE"] = (0 - $hash[$NUMASSU]["UC_CONTROLE"])
                $hash[$NUMASSU]["UC_MNTISU-UC_CONTROLE > 1"] = IsAbsoluteGapGreaterThan1 0 $hash[$NUMASSU]["UC_CONTROLE"]
            }

        } else { # add a new NUMASSU

            $NUMASSU_CONTROLE = $NUMASSU

            # compute euro related values if euro > 0
            if ( $EURO_CONTROLE -gt 0 ) {
                $EURO_DIFFERENCE_MNTISU_CONTROLE = ( 0 - $EURO_CONTROLE )
                $EURO_DIFFERENCE_MNTISU_CONTROLE_SUP1 = IsAbsoluteGapGreaterThan1 0 $EURO_CONTROLE
            } else {
                $EURO_DIFFERENCE_MNTISU_CONTROLE = 0
                $EURO_DIFFERENCE_MNTISU_CONTROLE_SUP1 = ""
            }
        
           
            # compute UC related values id uc > 0
            if ( $UC_CONTROLE -gt 0 ) {
                $UC_DIFFERENCE_MNTISU_CONTROLE = ( 0 - $UC_CONTROLE )
                $UC_DIFFERENCE_MNTISU_CONTROLE_SUP1 = IsAbsoluteGapGreaterThan1 0 $UC_CONTROLE
            } else {
                $UC_DIFFERENCE_MNTISU_CONTROLE = 0
                $UC_DIFFERENCE_MNTISU_CONTROLE_SUP1 = ""
            }
        
            # add key / value

            $hash2 = @{}

            $hash2.Add("NUMASSU", $NUMASSU)
            $hash2.Add("NUMASSU_CONTROLE", $NUMASSU)
            $hash2.Add("NUMASSU_PIVOT", "")
            $hash2.Add("€_MNTISU", 0)
            $hash2.Add("€_MNTSUP", 0)
            $hash2.Add("€_NBPENG", 0)
            $hash2.Add("€_CONTROLE", $EURO_CONTROLE)
            $hash2.Add("€_MNTISU-€_CONTROLE", $EURO_DIFFERENCE_MNTISU_CONTROLE)
            $hash2.Add("€_MNTISU-€_CONTROLE > 1", $EURO_DIFFERENCE_MNTISU_CONTROLE_SUP1)
            $hash2.Add("UC_MNTISU", 0)
            $hash2.Add("UC_MNTSUP", 0)
            $hash2.Add("UC_NBPENG", 0)
            $hash2.Add("UC_CONTROLE", $UC_CONTROLE)
            $hash2.Add("UC_MNTISU-UC_CONTROLE", $UC_DIFFERENCE_MNTISU_CONTROLE)
            $hash2.Add("UC_MNTISU-UC_CONTROLE > 1", $UC_DIFFERENCE_MNTISU_CONTROLE_SUP1)
            $hash2.Add("AbsentDuPivot", "VRAI")
            $hash2.Add("AbsentDuFichierDeControle", "")


            $hash.Add( $NUMASSU, $hash2)


        } 

       
    }

    "TOTAL PM €   : {0:N}" -f $TOTALPM_EURO | write-host  
    "TOTAL PM UC  : {0:N}" -f $TOTALPM_UC | write-host
    "TOTAL PM €+UC: {0:N}" -f ($TOTALPM_EURO + $TOTALPM_UC) | write-host
 
}


Function loadReferencePM {

    param (
        [string]$application,
        [string]$numlot,
        [hashtable]$hash
    )

    write-host "   START loadReferencePM"
   
    Measure-Command {


        # convert NUMLOT in pivot to value 1, 2 or 3

        $numlot = switch ( $numlot ) {

            "1" { "1" ; break }
            "2" { "2" ; break }
            "3" { "3" ; break }
            "20" { "1" ; break }
            "21" { "1" ; break }
            "22" { "2" ; break }
            "23" { "3" ; break }
            default { "*" }

        }


        $filter = "$refDir\{0}_{1}_*.csv" -f $application, $numlot

        foreach ( $file in Get-Item -Path $filter  ) {
        

            $fileNameSegments = $file.Name.split("_")
        
            $produit = $fileNameSegments[2]


            if ( $application -eq "NORA" ) {

                loadReferencePM_NORA $file.FullName $hash


            } else { # application = LOIC


                switch ( $produit ) {

                    "813" { loadReferencePM_LOIC_813 $file.FullName $hash }
                    "818" { loadReferencePM_LOIC_818 $file.FullName $hash }

                }                
            }
        }
    }
    
    write-host "   END loadReferencePM"     

}

# ##################################
# TEST
# ##################################

Function test_loadReferencePM {



    if ( !(get-variable "refDir" -ErrorAction SilentlyContinue) ) {
        
        $refDir = "$PSScriptRoot\..\..\ref"
    }


    [hashtable]$hash = @{}


    # test NORA lot 1
    loadReferencePM "NORA" "1" $hash

    # test LOIC lot 3
#    loadReferencePM "LOIC" "3" $hash


}


test_loadReferencePM




Function update_operations_counters2 {

    param ( 
    
        [hashtable]$hash_operations_counters2, 
        [hashtable]$hash_CDPRDTNORA2CDPRDT,
        [hashtable]$hash_referencePM,
        $operation
    )

#    write-host "   START update_operations_counters2"


        # build $key $CDPRDT;$CDCLLCTVT;$CDCNTRT;$NMCPT
                    
        $CDPRDT = $operation.NUMASSU.substring(0,3)

        # for NORA: convert code produit NORA into code produit CNP
        if ( $APPLICATION -eq "NORA" ) {

            Try {
                $CDPRDT = $hash_CDPRDTNORA2CDPRDT[ $CDPRDT ]
            } Catch {

                $msg = "Code produit NORA ({0}) non trouvé dans le fichier .\ref\CDPRDTNORA2CDPRDT.csv" -f $CDPRDT
                Write-Host $msg -f Red

            }
                         
        }
 
        $CDCLLCTVT = $operation.NUMCONTRAT.substring(0,5)
        $CDCNTRT =$operation.NUMCONTRAT.substring(0,10).substring(5,5)
        $NMCPT = "{0}{1}" -f $CDPRDT, $operation.NUMASSU.Substring(3)

        $key = "{0};{1};{2};{3};{4}" -f  $APPLICATION, $CDPRDT, $CDCLLCTVT, $CDCNTRT, $NMCPT

        if ( $hash_operations_counters2.ContainsKey( $key ) -eq $false ) {

            # hash_operations_counters2 fields
            $hash_operations_counters2_fields = @{}
            $hash_operations_counters2_fields.Add("FG-MNTISU",0)
            $hash_operations_counters2_fields.Add("FG-NBPENG",0)
            $hash_operations_counters2_fields.Add("UC-MNTISU",0)
            $hash_operations_counters2_fields.Add("UC-NBPENG",0)
            $hash_operations_counters2_fields.Add("MNTISU",0)
            $hash_operations_counters2_fields.Add("NBPENG",0)
            $hash_operations_counters2_fields.Add("MNTPM",0)
            $hash_operations_counters2_fields.Add("Ecart",0)
            $hash_operations_counters2_fields.Add("CpteAbsentDu","")
                
            $hash_operations_counters2.Add( $key, $hash_operations_counters2_fields )
        }


        $hash_operations_counters2[ $key ]["MNTISU"]+= [double]$operation.MNTISU.Replace(",",".")
        $hash_operations_counters2[ $key ]["NBPENG"]+= [double]$operation.NBPENG.Replace(",",".")


        if ( $operation.TYPSUP -eq "FG") {

            $hash_operations_counters2[ $key ][ "FG-MNTISU" ]+= [double]$operation.MNTISU.Replace(",",".")
            $hash_operations_counters2[ $key ][ "FG-NBPENG" ]+= [double]$operation.NBPENG.Replace(",",".")

        } else {

            $hash_operations_counters2[ $key ][ "UC-MNTISU" ]+= [double]$operation.MNTISU.Replace(",",".")
            $hash_operations_counters2[ $key ][ "UC-NBPENG" ]+= [double]$operation.NBPENG.Replace(",",".")
        }

        # look up for reference PM

        if ( $hash_referencePM.ContainsKey( $key  ) ) { # reference PM is found

            $hash_operations_counters2[ $key ]["MNTPM"] = $hash_referencePM[ $key ]

            # set reference PM to zero as used once
            $hash_referencePM[ $key ] = 0

        } else {  # reference PM is not found

            $hash_operations_counters2[ $key ]["CpteAbsentDu"] = "Fichier de contrôle"

        }



}                    

Function  add_accounts_with_referencePM_and_not_in_OPERATIONS_file  {

    param (

        [hashtable]$hash_referencePM,
        [hashtable]$hash_operations_counters2

    )

    write-host "   START add_accounts_with_referencePM_and_not_in_OPERATIONS_file"
   

    Measure-Command {
    
       
        foreach ( $account in $hash_referencePM.GetEnumerator() ) {

            if ( $hash_referencePM[ $account.key ] -gt 0 ) {


                # hash_operations_counters2 fields
                $hash_operations_counters2_fields = @{}
                $hash_operations_counters2_fields.Add("FG-MNTISU",0)
                $hash_operations_counters2_fields.Add("FG-NBPENG",0)
                $hash_operations_counters2_fields.Add("UC-MNTISU",0)
                $hash_operations_counters2_fields.Add("UC-NBPENG",0)
                $hash_operations_counters2_fields.Add("MNTISU",0)
                $hash_operations_counters2_fields.Add("NBPENG",0)
                $hash_operations_counters2_fields.Add("MNTPM",0)
                $hash_operations_counters2_fields.Add("Ecart",0)
                $hash_operations_counters2_fields.Add("CpteAbsentDu","")

                $hash_operations_counters2.Add( $account.key, $hash_operations_counters2_fields )

                $hash_operations_counters2[ $account.key ]["MNTPM"] = $hash_referencePM[ $account.key ]


                # compute gap
#                switch ( $APPLICATION ) {
#                    "NORA" { $hash_operations_counters2[ $key ]["Ecart"] = ( $hash_operations_counters2[ $key ]["NBPENG"] - $hash_operations_counters2[ $key ]["MNTPM"] ) ; break }
#                    "LOIC" { $hash_operations_counters2[ $key ]["Ecart"] = ( $hash_operations_counters2[ $key ]["MNTISU"] - $hash_operations_counters2[ $key ]["MNTPM"] ) ; break }
#                }


                $hash_operations_counters2[ $account.key ]["CpteAbsentDu"] = "Pivot"

            }

        }  





    }


    write-host "   END add_accounts_with_referencePM_and_not_in_OPERATIONS_file"
}

Function  print_operations_counters2  {

    param (

        [hashtable]$hash_operations_counters2

    )

    write-host "   START print_operations_counters2"

    $header = "sep=;`nAPPLICATION;CDPRDT;CDCLLCTVT;CDCNTRT;NMCPT;FG-MNTISU;FG-NBPENG;UC-MNTISU;UC-NBPENG;MNTISU;NBPENG;MNTPM;Ecart;CpteAbsentDu"
    
#    $outFile = "{0}\{1}" -f "C:\temp", ($zipFile.Name).Replace($zipFile.Extension,"-2.csv")

    $outFile = "{0}\{1}" -f $outDir, ($zipFile.Name).Replace($zipFile.Extension,"-2.csv")

    $header > $outFile

    Measure-Command {
    
        foreach ( $key in $hash_operations_counters2.Keys ) {

            $hash_operations_counters2[ $key ]["Ecart"] = $hash_operations_counters2[ $key ]["MNTISU"] - $hash_operations_counters2[ $key ]["MNTPM"]

            $line = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}" -f 
                $key,
                $hash_operations_counters2[ $key ]["FG-MNTISU"],
                $hash_operations_counters2[ $key ]["FG-NBPENG"],
                $hash_operations_counters2[ $key ]["UC-MNTISU"],
                $hash_operations_counters2[ $key ]["UC-NBPENG"],
                $hash_operations_counters2[ $key ]["MNTISU"],
                $hash_operations_counters2[ $key ]["NBPENG"],
                $hash_operations_counters2[ $key ]["MNTPM"],
                $hash_operations_counters2[ $key ]["Ecart"],
                $hash_operations_counters2[ $key ]["CpteAbsentDu"]

            $line >> $outFile

        }

    }


    write-host "   END print_operations_counters2"
}