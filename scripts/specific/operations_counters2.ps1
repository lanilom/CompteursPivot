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

                $msg = "Code produit NORA ({0}) non trouv� dans le fichier .\ref\CDPRDTNORA2CDPRDT.csv" -f $CDPRDT
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

            $hash_operations_counters2[ $key ]["CpteAbsentDu"] = "Fichier de contr�le"

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