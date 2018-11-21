# mountLetterDrive.ps1
#
# $result = mountLetterDrive "\\cyan.angers.cnp.fr\7-Fichiers Spéciaux\Y Pivot ACA\CompteursPivot"
#
# test if the mounting point (directory) is already associated to a drive letter -> return the drive letter followed by :. If not: search for the first free drive letter starting from A to Z.
# result:
# - failure: ""
# - success: drive letter to which the directory is mounted


function mountLetterDrive
{
    param (
     
        [string]$letterDriveMountingDirectory 

    )

    $driveLetter =""

    $availableDriveLetters = @( "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y","Z" )

    $usedDriveLetters = @{}




    $letterDriveMountingDirectory = $letterDriveMountingDirectory.Replace("Microsoft.PowerShell.Core\FileSystem::","")

    Get-PSDrive -PSProvider FileSystem | % {

        if ( $_.Root -eq $letterDriveMountingDirectory ) {

            if ( $driveLetter -eq "" ) {

#                $driveLetter = "{0}:" -f $_.Name
                 $driveLetter = $_.Name

            }

        } else {

            $usedDriveLetters.Add( $_.Name, "")

        }
    }

    # $directoy has to be mounted on a new drive
    if ( $driveLetter -eq "" ) {

        for ($i = 0 ; $i -lt $availableDriveLetters.Length ; $i++ ) {

            Try {

                $usedDriveLetters.Add( $availableDriveLetters[ $i ], "" )

                $driveLetter =  $availableDriveLetters[ $i ]

                Try {

                    $result = New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $letterDriveMountingDirectory -Scope Global

#                    $driveLetter = "{0}:" -f $driveLetter


                } Catch{

                    $Error

                    $driveLetter = ""

                }

                break

            } Catch {

                # the drive letter $availableDriveLetters[ $i ] is used

            }

        }

    }

    $driveLetter
}

