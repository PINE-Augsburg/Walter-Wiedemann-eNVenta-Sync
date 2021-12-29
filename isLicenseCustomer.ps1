$SQL_Server = 'WWSV05'
$SQL_DB = 'Wiedemann'

function isLicenseCustomer($kundenNr)
{
    $SQL_Query_Customer_License_Status = "SELECT C80_Licensing FROM KUNDEN WHERE KUNDENNR=$kundenNr"
    $licenseStatus = Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Query_Customer_License_Status

    if ($licenseStatus.Item(0) -eq '1')
    {
        #Write-Host "Lizenzkunde" -ForegroundColor Green
        return 1
    }
    else 
    {
        return 0
    }
}

function isLicenseArticle($customerNr, $articleNr, $licensingByCustomer)
{
    $SQL_Query_is_License_Article = "SELECT * FROM AdditionalItem WHERE ArticleID='$articleNr'"
    $isLicenseArticle = Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Query_is_License_Article
    
    $SQL_Query_Customer_Group = "SELECT GRUPPE FROM KUNDEN WHERE KUNDENNR='$customerNr'"
    $customerGroupNr = Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Query_Customer_Group

    $SQL_Query_is_Licensed_By_Customer = "SELECT * FROM I_LizenzKunden WHERE (Kundennr='$customerNr' OR Kundengruppe='$customerGroupNr') AND Verkaufsartikelnr='$articleNr'"
    $isLicensedByCustomer = Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Query_is_Licensed_By_Customer

    #Debug
    #Write-Host $customerGroupNr.Item('GRUPPE')
    #Write-Host $isLicenseArticle.Item('LizenzArtikelnr')
    if (-not [string]::IsNullOrEmpty($isLicenseArticle))
    {
        if ([string]::IsNullOrEmpty($isLicensedByCustomer))
        {
            return 1
        }
        else 
        {
            return 0
        }
    }
    else 
    {
        return 0
    }
    
}
