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

    $MNTPM_TOTAL = 0

    Import-Csv -Path $file.FullName -Delimiter ";" -Encoding Default  | % {

        $APPLICATION="NORA"
        $CDPRDT = $_."Code produit CNP"
        $CDCLLCTVT = ($_."Code collectivité").substring(0, 5)
        $CDCNTRT = $_."Code contrat collectif"
#        $NMCPT = "{0}{1}" -f $_."Code du produit (NORA)", ($_."N° de compte").substring(3, 8)
        $NMCPT = $_."N° de compte"
        $MNTPM =  [double]($_."Valeur de rachat fin N-1 (ind.)".Replace(",","."))

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

#[hashtable]$hash = @{}


# test NORA lot 1
#loadReferencePM "NORA" "1" $hash

# test LOIC lot 3
#loadReferencePM "LOIC" "3" $hash



