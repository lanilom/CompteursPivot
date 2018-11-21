
Function update_batch_counters {

    param (

        [hashtable] $hash_batchCounters,
        $operation
        
    )



    $hash_batchCounters[ "MNTISU" ]+= [double]$operation.MNTISU.Replace(",",".")
    $hash_batchCounters[ "MNTSUP" ]+= [double]$operation.MNTSUP.Replace(",",".")
    $hash_batchCounters[ "NBPENG" ]+= [double]$operation.NBPENG.Replace(",",".")


    if ( $_.TYPSUP -eq "FG") {
        $hash_batchCounters[ "FG-MNTISU" ]+= [double]$operation.MNTISU.Replace(",",".")
        $hash_batchCounters[ "FG-MNTSUP" ]+= [double]$operation.MNTSUP.Replace(",",".")
        $hash_batchCounters[ "FG-NBPENG" ]+= [double]$operation.NBPENG.Replace(",",".")
    } else {
        $hash_batchCounters[ "UC-MNTISU" ]+= [double]$operation.MNTISU.Replace(",",".")
        $hash_batchCounters[ "UC-MNTSUP" ]+= [double]$operation.MNTSUP.Replace(",",".")
        $hash_batchCounters[ "UC-NBPENG" ]+= [double]$operation.NBPENG.Replace(",",".")
    }


}


Function getReportHeader {

    $header = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12};{13};{14};{15};{16};{17};{18};{19};{20};{21};{22};{23};{24};{25};{26}" -f 
            "FICHIER",
            "NBLIGN-TOTAL", 
            "MNTISU", 
            "MNTSUP", 
            "NBPENG",
            "FG-MNTISU", 
            "FG-MNTSUP", 
            "FG-NBPENG",
            "UC-MNTISU", 
            "UC-MNTSUP", 
            "UC-NBPENG",
            "NBLIGN-ADHERENTES", 
            "NBLIGN-ARBITRAGE", 
            "NBLIGN-ASSURES", 
            "NBLIGN-COMMISSIONNEMENT", 
            "NBLIGN-CONTRAT", 
            "NBLIGN-COORDONNEES_BANCAIRES", 
            "NBLIGN-COTIS_CAT_PER", 
            "NBLIGN-ENVIRONNEMENT_FINANCIER", 
            "NBLIGN-FRAIS", 
            "NBLIGN-GARANTIES_ASSURES", 
            "NBLIGN-GARANTIES_CATPER", 
            "NBLIGN-GARANTIES_CONTRAT", 
            "NBLIGN-OPERATIONS", 
            "NBLIGN-OPTIONS_DE_RENTE", 
            "NBLIGN-PERSONNE_MORALE", 
            "NBLIGN-PERSONNE_PHYSIQUE"

    $header

}




Function print_batch_counters {

    param (

        [hashtable]$hash_batchCounters  

    )

    
    # print header if needed
        if ( (Test-Path -path "$outDir\report.csv") -eq $false ) { getReportHeader > "$outDir\report.csv" }


        # print data

        $line = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12};{13};{14};{15};{16};{17};{18};{19};{20};{21};{22};{23};{24};{25};{26}" -f 
            $zipFile.Name, 
            $hash_batchCounters[ "NBLIGN" ], 
            $hash_batchCounters[ "MNTISU" ], 
            $hash_batchCounters[ "MNTSUP" ], 
            $hash_batchCounters[ "NBPENG" ], 
            $hash_batchCounters[ "FG-MNTISU" ], 
            $hash_batchCounters[ "FG-MNTSUP" ], 
            $hash_batchCounters[ "FG-NBPENG" ], 
            $hash_batchCounters[ "UC-MNTISU" ], 
            $hash_batchCounters[ "UC-MNTSUP" ], 
            $hash_batchCounters[ "UC-NBPENG" ], 
            $hash_batchCounters[ "ADHERENTES" ], 
            $hash_batchCounters[ "ARBITRAGE" ], 
            $hash_batchCounters[ "ASSURES" ], 
            $hash_batchCounters[ "COMMISSIONNEMENT" ], 
            $hash_batchCounters[ "CONTRAT" ], 
            $hash_batchCounters[ "COORDONNEES_BANCAIRES" ], 
            $hash_batchCounters[ "COTIS_CAT_PER" ], 
            $hash_batchCounters[ "ENVIRONNEMENT_FINANCIER" ], 
            $hash_batchCounters[ "FRAIS" ], 
            $hash_batchCounters[ "GARANTIES_ASSURES" ], 
            $hash_batchCounters[ "GARANTIES_CATPER" ], 
            $hash_batchCounters[ "GARANTIES_CONTRAT" ], 
            $hash_batchCounters[ "OPERATIONS" ], 
            $hash_batchCounters[ "OPTIONS_DE_RENTE" ], 
            $hash_batchCounters[ "PERSONNE_MORALE" ], 
            $hash_batchCounters[ "PERSONNE_PHYSIQUE" ] 
        
        $line >> "$outDir\report.csv"
}