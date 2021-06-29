############################################################################################################
#
# SCRIPT       : Export-ADAccounts.ps1
#
# DESCRIPTION  : Ce script permet de faire une extraction des comptes utilisateurs et/ou ordinateurs
#                de l'Active Directory.
#
# AUTEUR       : Johann BARON
#
# DATE         : 17/11/2017
#
# MODIFICATIONS :
#
############################################################################################################


# -----------------------------------------
# Paramétrage de l'affichage de la console
# -----------------------------------------

[Console]::ForegroundColor = "Green"
[Console]::BackgroundColor = "Black"
[Console]::WindowHeight = 50
[Console]::BufferHeight = 50
[Console]::WindowWidth = 100
[Console]::BufferWidth = 100
Clear-Host


# ---------------------------
# Fonction du menu principal
# ---------------------------

Function MainMenu {

    Clear-Host
    ""
    "           *******************************************"
    "           *                                         *"
    "           *           Export des comptes            *"
    "           *         de l'Active Directory           *"
    "           *                                         *"
    "           *******************************************"
    ""
    ""
    ""
    "               xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    "               x                                x"
    "               x         Type de compte         x"
    "               x         --------------         x"
    "               x                                x"
    "               x   1) Comptes utilisateurs      x"
    "               x   2) Comptes ordinateurs       x"
    "               x                                x"
    "               x   Q) Quitter                   x"
    "               x                                x"
    "               xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    ""
    ""
    ""
    [Console]::ForegroundColor = "White"

    $Choix = Read-Host "   Choix "

    [Console]::ForegroundColor = "Green"

    if($Choix -eq 1){
        Write-Host "
     L'export des comptes utilisateurs de l'Active Directory est en cours.
     Merci de patienter..." -ForegroundColor Cyan
        Start-Sleep -Seconds 2
        GetADUsers

    }elseif($Choix -eq 2){
        Write-Host "
     L'export des comptes ordinateurs de l'Active Directory est en cours.
     Merci de patienter..." -ForegroundColor Cyan
        Start-Sleep -Seconds 2
        GetADComputers

    }elseif($Choix -like "q"){
        Exit

    }else{
        Write-Host "
         Entrée invalide !" -ForegroundColor Red
        Start-Sleep -Seconds 2
        MainMenu
    }

}


# -----------------------------------------------
# Fonction d'export des comptes utilisateurs
# -----------------------------------------------

Function GetADUsers{

    $CurrentDate = (Get-Date -UFormat "%Y%m%d_%H%M%S")

    $CurrentLocation = (Get-Location)

    $ExportFileName = "$CurrentLocation\Export_ADUsers_$CurrentDate.csv"

    $ADName = ""

    Get-ADUser -Filter * -SearchBase $ADName -Properties * |     select -Property SamAccountName,CN,comment,    @{Name="Lecteurs Réseaux";Expression={($_ | select -ExpandProperty NetworkDrive) -join "`n`r" -replace ";",":"}},
    Description,Enabled,Modified | Sort-Object -Property @{Expression = "CN"; Descending = $false} | 
    Export-Csv $ExportFileName -Delimiter ";" -NoTypeInformation -Encoding UTF8

    Write-Host ""    Write-Host ""    Write-Host "   Export terminé !" -ForegroundColor Yellow    Write-Host ""    Write-Host " Le fichier est enregistré sous $ExportFileName"    Write-Host ""    Write-Host ""    Write-Host "    ---> Appuyer sur une touche pour continuer <---" -ForegroundColor White    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")    MainMenu}


# ----------------------------------------------
# Fonction d'export des comptes ordinateurs
# ----------------------------------------------

Function GetADComputers{

    $CurrentDate = (Get-Date -UFormat "%Y%m%d_%H%M%S")

    $CurrentLocation = (Get-Location)

    $ExportFileName = "$CurrentLocation\Export_ADComputers_$CurrentDate.csv"

    Get-ADComputer -Filter * -SearchBase $ADName -Properties Name,Description,
    LastLogonDate | select -Property Name,Description,LastLogonDate | 
    Sort-Object -Property @{Expression = "Name"; Descending = $false} | 
    Export-Csv $ExportFileName -Delimiter ";" -NoTypeInformation -Encoding UTF8

    Write-Host ""    Write-Host ""    Write-Host "   Export terminé !" -ForegroundColor Yellow    Write-Host ""    Write-Host " Le fichier est enregistré sous $ExportFileName"    Write-Host ""    Write-Host ""    Write-Host "    ---> Appuyer sur une touche pour continuer <---" -ForegroundColor White    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")    MainMenu}


MainMenu