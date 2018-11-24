
# *****************************************************
# load XML config file
# *****************************************************

Function load_XMLCONFIG {
    param (
        [string]$configDir,
        [string]$logDir
    )

    $XMLCONFIG = $null

    Try {
        [xml]$XMLCONFIG = Get-Content -Path "$configDir\config.xml" -Encoding UTF8
    } Catch {
        $msg = "ERREUR: Le fichier de configuration XML $configDir\config.xml n a pu être chargé. Merci de corriger l erreur de syntaxe indiquée dans le fichier {0} et de relancer le traitement.`n" -f "$logDir\log.txt"
        write-host $msg -f Red
        $msg >> "$logDir\log.txt"
        $error >> "$logDir\log.txt"
        exit
    }

    $XMLCONFIG

}