


Function update_operations_counters {
    
    param (

        [hashtable]$hash_operations_counters,
        $operation
    )

    # OPERATIONS counters

    # NBPENG: nombre de parts
    # VALTIT: valeur liquidative
    # MNTISU: Montant net de l'opération compte épargne réparti sur le support
    # MNTSUP: Montant de l'opération sur le support net de frais titre

    # if [TYPSUP] = UC then [MNTISU] = [MNTSUP] = [NBPENG] * [VALTIT]
    # if [TYPSUP] = UC then [MNTISU] = [MNTSUP] = [NBPENG]


    $contratColl = $operation.NUMCONTRAT

                
    if ( $hash_operations_counters.ContainsKey( $contratColl ) -eq $false ) {
                   
        $hash_operations_counters.Add( $contratColl, @{} )
    }


    $hash_operations_counters[ $contratColl ]["NBLIGN"] += 1

    $hash_operations_counters[ $contratColl ]["MNTISU"]+= [double]$operation.MNTISU.Replace(",",".")
    $hash_operations_counters[ $contratColl ]["MNTSUP"]+= [double]$operation.MNTSUP.Replace(",",".")
    $hash_operations_counters[ $contratColl ]["NBPENG"]+= [double]$operation.NBPENG.Replace(",",".")


    if ( $operation.TYPSUP -eq "FG") {
        $hash_operations_counters[ $contratColl ][ "FG-MNTISU" ]+= [double]$operation.MNTISU.Replace(",",".")
        $hash_operations_counters[ $contratColl ][ "FG-MNTSUP" ]+= [double]$operation.MNTSUP.Replace(",",".")
        $hash_operations_counters[ $contratColl ][ "FG-NBPENG" ]+= [double]$operation.NBPENG.Replace(",",".")
    } else {
      $hash_operations_counters[ $contratColl ][ "UC-MNTISU" ]+= [double]$operation.MNTISU.Replace(",",".")
      $hash_operations_counters[ $contratColl ][ "UC-MNTSUP" ]+= [double]$operation.MNTSUP.Replace(",",".")
      $hash_operations_counters[ $contratColl ][ "UC-NBPENG" ]+= [double]$operation.NBPENG.Replace(",",".")
   }


}


Function print_operations_counters {

    param ( 
        [hashtable]$hash_operations_counters 
    )
    
    
        $header = "CONTRATCOLL;NBLIGN;MNTISU;MNTSUP;NBPENG;FG-MNTISU;FG-MNTSUP;FG-NBPENG;UC-MNTISU;UC-MNTSUP;UC-NBPENG"
    
        $outFile = "{0}\{1}" -f $outDir, ($zipFile.Name).Replace($zipFile.Extension,".csv")

        $header > $outFile

        foreach ( $contratColl in $hash_operations_counters.Keys ) {

            $line = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f 
                $contratColl,
                $hash_operations_counters[ $contratColl ]["NBLIGN"],
                $hash_operations_counters[ $contratColl ]["MNTISU"],
                $hash_operations_counters[ $contratColl ]["MNTSUP"],
                $hash_operations_counters[ $contratColl ]["NBPENG"],
                $hash_operations_counters[ $contratColl ]["FG-MNTISU"],
                $hash_operations_counters[ $contratColl ]["FG-MNTSUP"],
                $hash_operations_counters[ $contratColl ]["FG-NBPENG"],
                $hash_operations_counters[ $contratColl ]["UC-MNTISU"],
                $hash_operations_counters[ $contratColl ]["UC-MNTSUP"],
                $hash_operations_counters[ $contratColl ]["UC-NBPENG"]

            $line >> $outFile
        }    
    
    
     

}