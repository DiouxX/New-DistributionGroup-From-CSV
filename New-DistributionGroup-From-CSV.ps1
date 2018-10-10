<#
.SYNOPSIS
    Ce script permet de créer des listes de distribution avec leurs membres depuis un fichier CSV
.DESCRIPTION
    Ce script permet de créer un ou plusieurs liste de diffusion et d'y assigner différents membres selon un fichier CSV
.PARAMETER PathCsv
    Ce paramètre est obligatoire, il permet de spécifier le chemin du fichier CSV
.PARAMETER HeaderDistributionGroup
    Ce paramètre est obligatoire, il permet de spécifier le header du fichier CSV reprenant le nom du groupe de distribution
.PARAMETER HeaderMailAddress
    Ce paramètre est obligatoire, il permet de spécifier le header du fichier CSV reprenant les adresses mail des utilisateurs à ajouter au groupe de distribution
.PARAMETER OrganizationalUnit
    Ce paramètre est optionel, il permet de spécifier l'OU AD pour la création des groupe de distribution
.PARAMETER Delimiter
    Ce paramètre est optionnel, il permet de spécifier le séparateur de champs du fichier CSV passé en paramètre [Defaut: ;]
.EXAMPLE
C:PS> .\New-DistributionGroup-From-CSV.ps1 -PathCsv .\csv-distributionGroup.csv -HeaderDistributionGroup DistributionGroup -HeaderMailAddress Mail
.EXAMPLE
C:PS> .\New-DistributionGroup-From-CSV.ps1 -PathCsv .\csv-distributionGroup.csv -HeaderDistributionGroup DistributionGroup -HeaderMailAddress Mail -Delimiter ","
.EXAMPLE
C:PS> .\New-DistributionGroup-From-CSV.ps1 -PathCsv .\csv-distributionGroup.csv -HeaderDistributionGroup DistributionGroup -HeaderMailAddress Mail -OrganizationalUnit "DL from GW" -Delimiter ","

#>

Param(
    [Parameter(Mandatory=$true)]
    [String]$PathCsv,
    [Parameter(Mandatory=$true)]
    [String]$HeaderDistributionGroup,
    [Parameter(Mandatory=$true)]
    [String]$HeaderMailAddress,
    [Parameter(Mandatory=$false)]
    [String]$OrganizationalUnit,
    [Parameter(Mandatory=$false)]
    [String]$Delimiter
)


#Attribution du delimteur ";" par défaut
if( !$Delimiter )
{
    $Delimiter=";"
}

$fileCSV=Import-Csv $PathCsv -Delimiter $Delimiter

#Fonction pour ajouter l'utilisateur au groupe de distribution
function Add-Member{
    Param(
    [String]$DistributionGroup,
    [String]$Mail
    )
    
    Try
    {
        Add-DistributionGroupMember -Identity $DistributionGroup -Member $Mail -ErrorAction Stop
        Write-Host -ForegroundColor Green "Ajout de $Mail au groupe de distribution $DistributionGroup`:`t[OK]"
    }
    Catch [System.Exception]
    {
        if($_.FullyQualifiedErrorId -match 'AlreadyExists')
        {
            Write-Host -ForegroundColor Yellow "$Mail déja présent dans le groupe de distribution $DistributionGroup"
        }
        else
        {
            Write-Output "`n Exception :`n "$_.Exception
            Write-Output "`n FullyQualifiedErrorId :`n"  $_.FullyQualifiedErrorId 
        }
    }
}

foreach($item in $fileCSV)
{
    #Si le groupe de distribution existe => Ajout de l'utilisateur au groupe existant
    if(Get-DistributionGroup $item.$HeaderDistributionGroup -ErrorAction 'SilentlyContinue')
    {
        Write-Host -ForegroundColor Yellow "`nLe groupe de diffusion" $item.$HeaderDistributionGroup "existe"
        Add-Member -DistributionGroup $item.$HeaderDistributionGroup -Mail $item.$HeaderMailAddress
    }
    #Si le groupe n'existe pas => Création du groupe et après ajout de l'utilisateur au groupe
    else
    {
        Write-Host -ForegroundColor Yellow "`nLe groupe de diffusion" $item.$HeaderDistributionGroup "n'existe pas"
        Write-Host -ForegroundColor Yellow "Création du groupe de diffusion "$item.$HeaderDistributionGroup "..."

        Try
        {
            if(!$OrganizationalUnit)
            {
                New-DistributionGroup -Name $item.$HeaderDistributionGroup | Out-Null
                Write-Host -ForegroundColor Green "Création du Groupe de distribution "$item.$HeaderDistributionGroup " :`t[OK]"
            }
            else
            {
                New-DistributionGroup -Name $item.$HeaderDistributionGroup -OrganizationalUnit $OrganizationalUnit | Out-Null
                Write-Host -ForegroundColor Green "Création du Groupe de distribution "$item.$HeaderDistributionGroup " dans l'OU $OrganizationalUnit :`t[OK]"
            }
            #Sleep pour synchro entre AD
            Write-Host -ForegroundColor Yellow "`nSynchronisation entre AD. Attendre 30 secondes..."
            Start-Sleep -s 30
            Add-Member -DistributionGroup $item.$HeaderDistributionGroup -Mail $item.$HeaderMailAddress
        }
        Catch
        {
            Write-Host -ForegroundColor Red "`nCréation du Groupe de distribution "$item.$HeaderDistributionGroup " :`t[KO]"
            Write-Output "`t $_.Exception"
        }
    }
}
