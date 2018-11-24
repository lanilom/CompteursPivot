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

Function initHashFields {

    $hash2 = @{}

    $hash2.Add("NUMASSU_CONTROLE", "")
    $hash2.Add("NUMASSU_PIVOT", "")

    $hash2.Add("EURO_MNTISU", 0)
    $hash2.Add("EURO_MNTSUP", 0)
    $hash2.Add("EURO_NBPENG", 0)
    $hash2.Add("EURO_CONTROLE", 0)

    $hash2.Add("UC_MNTISU", 0)
    $hash2.Add("UC_MNTSUP", 0)
    $hash2.Add("UC_NBPENG", 0)
    $hash2.Add("UC_CONTROLE", 0)


    $hash2

}

Function loadReferencePM_LOIC_813 {

    param (
        $fileFullname,
        [hashtable]$hash

    )

   write-host "START loadReferencePM_LOIC_813"

    $TOTALPM_EURO = 0
    $TOTALPM_UC = 0

    Import-Csv -Path $file.FullName -Delimiter ";" -Encoding Default  | % {
     
        $NUMASSU = $_."Ident compte LOIC"

        # get EURO and UC values

        # EURO 
        $EURO_CONTROLE = [double]($_."Epargne au 31/12/n-1".Replace(",","."))
         

        # UC
        $UC_CONTROLE = 0


        # compute total PM
        $TOTALPM_EURO += $EURO_CONTROLE
        $TOTALPM_UC += $UC_CONTROLE


        # search if NUMASSU was already inserted in the index


        if ( $hash.ContainsKey( $NUMASSU ) ) { # update NUMASSU € and UC related values

            if ( $EURO_CONTROLE -gt 0 ) {

                $hash[$NUMASSU]["EURO_CONTROLE"] += $EURO_CONTROLE

            }

            if ( $UC_CONTROLE -gt 0 ) {
                $hash[$NUMASSU]["UC_CONTROLE"] += $UC_CONTROLE

            }

        } else { # add a new NUMASSU
            
            $hash.Add( $NUMASSU, (initHashFields) )

            $hash[$NUMASSU]["NUMASSU_CONTROLE"] = $NUMASSU
            $hash[$NUMASSU]["EURO_CONTROLE"] = $EURO_CONTROLE
            $hash[$NUMASSU]["UC_CONTROLE"] = $UC_CONTROLE


        } 

       
    }

    "TOTAL PM EURO   : {0:N}" -f $TOTALPM_EURO | write-host  
    "TOTAL PM UC     : {0:N}" -f $TOTALPM_UC | write-host
    "TOTAL PM EURO+UC: {0:N}" -f ($TOTALPM_EURO + $TOTALPM_UC) | write-host

    write-host "END loadReferencePM_LOIC_813"


}

Function loadReferencePM_LOIC_818 {

    param (
        $fileFullname,
        [hashtable]$hash

    )

    write-host "START loadReferencePM_LOIC_818"


    $TOTALPM_EURO = 0
    $TOTALPM_UC = 0

    Import-Csv -Path $file.FullName -Delimiter ";" -Encoding Default  | % {
     
        $NUMASSU = $_."Ident compte LOIC"

        # get EURO and UC values

        # EURO
        $EURO_PART_ENTREPRISE = [double]($_."Epagne part entreprise".Replace(",","."))
        $EURO_PART_SALARIE = [double]($_."Epagne part salarié".Replace(",","."))
        $EURO_CONTROLE =  $EURO_PART_ENTREPRISE + $EURO_PART_SALARIE
         

        # UC
        $UC_CONTROLE = 0


        # compute total PM
        $TOTALPM_EURO += $EURO_CONTROLE
        $TOTALPM_UC += $UC_CONTROLE


        # search if NUMASSU was already inserted in the index


        if ( $hash.ContainsKey( $NUMASSU ) ) { # update NUMASSU € and UC related values

            if ( $EURO_CONTROLE -gt 0 ) {

                $hash[$NUMASSU]["EURO_CONTROLE"] += $EURO_CONTROLE

            }

            if ( $UC_CONTROLE -gt 0 ) {
                $hash[$NUMASSU]["UC_CONTROLE"] += $UC_CONTROLE

            }

        } else { # add a new NUMASSU
            
            $hash.Add( $NUMASSU, (initHashFields) )

            $hash[$NUMASSU]["NUMASSU_CONTROLE"] = $NUMASSU
            $hash[$NUMASSU]["EURO_CONTROLE"] = $EURO_CONTROLE
            $hash[$NUMASSU]["UC_CONTROLE"] = $UC_CONTROLE


        } 

       
    }

    "TOTAL PM EURO   : {0:N}" -f $TOTALPM_EURO | write-host  
    "TOTAL PM UC     : {0:N}" -f $TOTALPM_UC | write-host
    "TOTAL PM EURO+UC: {0:N}" -f ($TOTALPM_EURO + $TOTALPM_UC) | write-host



    

    write-host "END loadReferencePM_LOIC_818"

}

Function loadReferencePM_NORA {

    param (
        $fileFullname,
        [hashtable]$hash

    )

   write-host "START loadReferencePM_NORA"

    $TOTALPM_EURO = 0
    $TOTALPM_UC = 0



    $CDPRDT_NORA_FIELDLABEL = setConfigFieldLabel $XMLCONFIG "NORA" "ALL" "CDPRDT_NORA"
    $test1 = doesfieldLabelBelongToCSVFileHeader $CDPRDT_NORA_FIELDLABEL ";" $fileFullname

    $NMCPT_FIELDLABEL = setConfigFieldLabel $XMLCONFIG "NORA" "ALL" "NMCPT"
    $test2 = doesfieldLabelBelongToCSVFileHeader $NMCPT_FIELDLABEL ";" $fileFullname
    
    $EURO_CONTROLE_FIELDLABEL = setConfigFieldLabel $XMLCONFIG "NORA" "ALL" "EURO_CONTROLE"
    $test3 = doesfieldLabelBelongToCSVFileHeader $EURO_CONTROLE_FIELDLABEL ";" $fileFullname
    
    $UC_CONTROLE_FIELDLABEL = setConfigFieldLabel $XMLCONFIG "NORA" "ALL" "UC_CONTROLE"
    $test4 = doesfieldLabelBelongToCSVFileHeader $UC_CONTROLE_FIELDLABEL ";" $fileFullname

    # exit if one field is not prese in the header
    if ( ($test1 -ne $true) -or ($test2 -ne $true) -or ($test3 -ne $true) -or ($test4 -ne $true) ) {

        exit
    }


    Import-Csv -Path $file.FullName -Delimiter ";" -Encoding Default  | % {

        # compute NUMASSU

#        $CDPRDT_NORA = $_."Code du produit (NORA)"
        $CDPRDT_NORA = $_.$CDPRDT_NORA_FIELDLABEL

#        $NMCPT = $_."N° de compte"
        $NMCPT = $_.$NMCPT_FIELDLABEL
        $NMCPT = $NMCPT.Substring(3,8)
        $NUMASSU = "{0}{1}" -f $CDPRDT_NORA, $NMCPT

        # get EURO and UC values

        # EURO = [Valeur de rachat fin N-1 (ind.)]
#        $EURO_CONTROLE = [double]($_."Valeur de rachat fin N-1 (ind.)".Replace(",","."))
        $EURO_CONTROLE = [double]($_.$EURO_CONTROLE_FIELDLABEL.Replace(",","."))

        # UC = [Valorisation de la garantie UC au 31/12/AA-1]
#        $UC_CONTROLE = [double]($_."Valorisation de la garantie UC au 31/12/AA-1".Replace(",","."))
        $UC_CONTROLE = [double]($_.$UC_CONTROLE_FIELDLABEL.Replace(",","."))


        # compute total PM
        $TOTALPM_EURO += $EURO_CONTROLE
        $TOTALPM_UC += $UC_CONTROLE


        # search if NUMASSU was already inserted in the index


        if ( $hash.ContainsKey( $NUMASSU ) ) { # update NUMASSU € and UC related values

            if ( $EURO_CONTROLE -gt 0 ) {

                $hash[$NUMASSU]["EURO_CONTROLE"] += $EURO_CONTROLE

            }

            if ( $UC_CONTROLE -gt 0 ) {
                $hash[$NUMASSU]["UC_CONTROLE"] += $UC_CONTROLE

            }

        } else { # add a new NUMASSU
            
            $hash.Add( $NUMASSU, (initHashFields) )

            $hash[$NUMASSU]["NUMASSU_CONTROLE"] = $NUMASSU
            $hash[$NUMASSU]["EURO_CONTROLE"] = $EURO_CONTROLE
            $hash[$NUMASSU]["UC_CONTROLE"] = $UC_CONTROLE


        } 

       
    }

    "TOTAL PM EURO   : {0:N}" -f $TOTALPM_EURO | write-host  
    "TOTAL PM UC     : {0:N}" -f $TOTALPM_UC | write-host
    "TOTAL PM EURO+UC: {0:N}" -f ($TOTALPM_EURO + $TOTALPM_UC) | write-host

    write-host "END loadReferencePM_NORA"
 
}

Function loadReferencePM {

    param (
        [string]$application,
        [string]$numlot,
        [hashtable]$hash
    )

    write-host "START loadReferencePM"
   
    


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
    
    
    write-host "END loadReferencePM"     

}

Function loadOperationsPM {

    param (
        [string]$operationsFileFullName,
        [hashtable]$hash
    )

    write-host "START loadOperationsPM"


    $TOTALPM_EURO = 0
    $TOTALPM_UC = 0


    Import-Csv -Path $operationsFileFullName -Delimiter ";" | % {

        $NUMASSU = $_.NUMASSU
        $TYPSUP = $_.TYPSUP
        
        $MNTISU = [double]($_.MNTISU.Replace(",","."))
        $MNTSUP = [double]($_.MNTSUP.Replace(",","."))
        $NBPENG = [double]($_.NBPENG.Replace(",","."))

        # search if NUMASSU was already inserted in the index

        if ( $hash.ContainsKey( $NUMASSU ) ) { # update NUMASSU € and UC related values

            $hash[$NUMASSU]["NUMASSU_PIVOT"] = $NUMASSU


            if ( ( $TYPSUP -eq "FG" ) -and ( $MNTISU -gt 0 ) ) {

                $hash[$NUMASSU]["EURO_MNTISU"] += $MNTISU

#                $hash[$NUMASSU]["EURO_MNTISU-EURO_CONTROLE"] = ( $hash[$NUMASSU]["EURO_MNTISU"] - $hash[$NUMASSU]["EURO_CONTROLE"] )
#                $hash[$NUMASSU]["EURO_MNTISU-EURO_CONTROLE > 1"] = IsAbsoluteGapGreaterThan1 $hash[$NUMASSU]["EURO_MNTISU"] $hash[$NUMASSU]["EURO_CONTROLE"]

                $hash[$NUMASSU]["EURO_MNTSUP"] += $MNTSUP
                $hash[$NUMASSU]["EURO_NBPENG"] += $NBPENG

               # compute total PM
                $TOTALPM_EURO += $MNTISU
      
            }

            if ( ( $TYPSUP -eq "UC" ) -and ( $MNTISU -gt 0 ) ) {

                $hash[$NUMASSU]["UC_MNTISU"] += $MNTISU

#                $hash[$NUMASSU]["UC_MNTISU-UC_CONTROLE"] = ( $hash[$NUMASSU]["UC_MNTISU"] - $hash[$NUMASSU]["UC_CONTROLE"] )
#                $hash[$NUMASSU]["UC_MNTISU-UC_CONTROLE > 1"] = IsAbsoluteGapGreaterThan1 $hash[$NUMASSU]["UC_MNTISU"] $hash[$NUMASSU]["UC_CONTROLE"]

                $hash[$NUMASSU]["UC_MNTSUP"] += $MNTSUP
                $hash[$NUMASSU]["UC_NBPENG"] += $NBPENG

                # compute total PM
                $TOTALPM_UC += $MNTISU

            }


#            $hash[$NUMASSU]["AbsentDuPivot"] = ""


        } else { # add a new NUMASSU

            # add key / value

            $hash.Add( $NUMASSU, (initHashFields) )

#            $hash[$NUMASSU]["NUMASSU"] = $NUMASSU
#            $hash[$NUMASSU]["NUMASSU_CONTROLE"] = ""
            $hash[$NUMASSU]["NUMASSU_PIVOT"] = $NUMASSU

            # euro related values

#            $hash[$NUMASSU]["EURO_MNTISU"] = 0
#            $hash[$NUMASSU]["EURO_MNTSUP"] = 0
#            $hash[$NUMASSU]["EURO_NBPENG"] = 0
#            $hash[$NUMASSU]["EURO_CONTROLE"] = 0

#            $hash[$NUMASSU]["EURO_MNTISU-EURO_CONTROLE"] = 0
#            $hash[$NUMASSU]["EURO_MNTISU-EURO_CONTROLE > 1"] = ""


            if ( ( $TYPSUP -eq "FG" ) -and ( $MNTISU -gt 0 ) ) {

                $hash[$NUMASSU]["EURO_MNTISU"] = $MNTISU

#                $hash[$NUMASSU]["EURO_MNTISU-EURO_CONTROLE"] = ( $hash[$NUMASSU]["EURO_MNTISU"] - $hash[$NUMASSU]["EURO_CONTROLE"] )
#                $hash[$NUMASSU]["EURO_MNTISU-EURO_CONTROLE > 1"] = IsAbsoluteGapGreaterThan1 $hash[$NUMASSU]["EURO_MNTISU"] $hash[$NUMASSU]["EURO_CONTROLE"]

                $hash[$NUMASSU]["EURO_MNTSUP"] = $MNTSUP
                $hash[$NUMASSU]["EURO_NBPENG"] = $NBPENG


                # compute total PM
                $TOTALPM_EURO += $MNTISU

                
            }



            # UC related values

#            $hash[$NUMASSU]["UC_MNTISU"] = 0
#            $hash[$NUMASSU]["UC_MNTSUP"] = 0
#            $hash[$NUMASSU]["UC_NBPENG"] = 0
#            $hash[$NUMASSU]["UC_CONTROLE"] = 0

#            $hash[$NUMASSU]["UC_MNTISU-UC_CONTROLE"] = 0
#            $hash[$NUMASSU]["UC_MNTISU-UC_CONTROLE > 1"] = ""



            if ( ( $TYPSUP -eq "UC" ) -and ( $MNTISU -gt 0 ) ) {

                $hash[$NUMASSU]["UC_MNTISU"] = $MNTISU

#                $hash[$NUMASSU]["UC_MNTISU-UC_CONTROLE"] = ( $hash[$NUMASSU]["UC_MNTISU"] - $hash[$NUMASSU]["UC_CONTROLE"] )
#                $hash[$NUMASSU]["UC_MNTISU-UC_CONTROLE > 1"] = IsAbsoluteGapGreaterThan1 $hash[$NUMASSU]["UC_MNTISU"] $hash[$NUMASSU]["UC_CONTROLE"]

                $hash[$NUMASSU]["UC_MNTSUP"] = $MNTSUP
                $hash[$NUMASSU]["UC_NBPENG"] = $NBPENG


                # compute total PM
                $TOTALPM_UC += $MNTISU

            }

#            $hash[$NUMASSU]["AbsentDuPivot"] = ""
#            $hash[$NUMASSU]["AbsentDuFichierDeControle"] = "VRAI"


        }
    }

    "TOTAL PM EURO   : {0:N}" -f $TOTALPM_EURO | write-host  
    "TOTAL PM UC     : {0:N}" -f $TOTALPM_UC | write-host
    "TOTAL PM EURO+UC: {0:N}" -f ($TOTALPM_EURO + $TOTALPM_UC) | write-host

    write-host "END loadOperationsPM"

}

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

    write-host "START add_accounts_with_referencePM_and_not_in_OPERATIONS_file"
   

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


    write-host "END add_accounts_with_referencePM_and_not_in_OPERATIONS_file"
}

Function isEmpty {
    param ( [string]$value="" )

    $return = "VRAI"

    if ( $value -ne "" ) {

        $return = ""

    }

    $return
}

Function IsValueGreaterThan1 {

    param (
        [double]$value=0

    )

    $return = ""


    if ( [math]::Abs( $value ) -gt 1 ) {

        $return = "VRAI"

    } 

    $return

}

Function  print_operations_counters2  {

    param (
        $APPLICATION,
        $NUMLOT,
        [hashtable]$hash

    )

    write-host "START print_operations_counters2"

    $header = "sep=;`nAPPLICATION;NUMLOT;NUMASSU;EURO_MNTISU;EURO_MNTSUP;EURO_NBPENG;EURO_CONTROLE;EURO_MNTISU-EURO_CONTROLE;EURO_MNTISU-EURO_CONTROLE > 1;UC_MNTISU;UC_MNTSUP;UC_NBPENG;UC_CONTROLE;UC_MNTISU-UC_CONTROLE;UC_MNTISU-UC_CONTROLE > 1;AbsentDuPivot;AbsentDuFichierDeControle"
    
    $outFile = "{0}\{1}" -f $outDir, ($zipFile.Name).Replace($zipFile.Extension,"-2.csv")

#$outFile = "c:\temp\test.csv"

    $header > $outFile

#    $lines=@()

#    $lines += $header 
    
        foreach ( $NUMASSU in $hash.Keys ) {


            $euro_gap = ( $hash[ $NUMASSU ]["EURO_MNTISU"] - $hash[ $NUMASSU ]["EURO_CONTROLE"] )
            $euro_gap_greater_1 =  IsValueGreaterThan1 $euro_gap 

            $uc_gap =  ( $hash[ $NUMASSU ]["UC_MNTISU"] - $hash[ $NUMASSU ]["UC_CONTROLE"] )
            $uc_gap_greater_1 = IsValueGreaterThan1 $uc_gap


            $isEmptyNUMASSU_PIVOT = isEmpty ( $hash[$NUMASSU]["NUMASSU_PIVOT"] )
            $isEmptyNUMASSU_CONTROLE = isEmpty ( $hash[$NUMASSU]["NUMASSU_CONTROLE"] )

           $line = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12};{13};{14};{15};{16}" -f 
                $APPLICATION,
                $NUMLOT,
                $NUMASSU,
                $hash[ $NUMASSU ]["EURO_MNTISU"],
                $hash[ $NUMASSU ]["EURO_MNTSUP"],
                $hash[ $NUMASSU ]["EURO_NBPENG"],
                $hash[ $NUMASSU ]["EURO_CONTROLE"],
                $euro_gap,
                $euro_gap_greater_1,
                $hash[ $NUMASSU ]["UC_MNTISU"],
                $hash[ $NUMASSU ]["UC_MNTSUP"],
                $hash[ $NUMASSU ]["UC_NBPENG"],
                $hash[ $NUMASSU ]["UC_CONTROLE"],
                $uc_gap,
                $uc_gap_greater_1,
                $isEmptyNUMASSU_PIVOT,
                $isEmptyNUMASSU_CONTROLE 
                
           $line >> $outfile

        }

#        $lines > $outfile

    


    write-host "END print_operations_counters2"
}

Function setConfigFieldLabel {
    param (
        $XMLCONFIG,
        $application,
        $produit,
        $fieldCode
    
    )
 
    $label=""

    $xmlconfig.SelectNodes('//config/fields/field') | % {

        if ( ( $_.application -eq $application ) -and ( $_.produit -eq $produit ) -and ( $_.code -eq $fieldCode ) ) {
            
            $label = $_.label
        }

#        $msg = '{0,-5} {1,-3 } {2,-25} {3,-35}' -f $_.application, $_.produit, $_.code, $_.label 

#        Write-host $msg

#        $key = "{0};{1};{2}" -f $_.application, $_.produit, $_.code

#        $hash_fields.Add( $key, $_.label)

    }

    $label
}

Function testConfigFields {

    


}
