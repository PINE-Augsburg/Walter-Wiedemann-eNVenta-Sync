

$URLLicensePosition = 'http://api.production.app-project2028.wiedemann-augsburg.de/api/v1/license-articles/?page=1&paginateBy=20&articleNumbers='

$URLLicensePositionData = 'http://api.production.app-project2028.wiedemann-augsburg.de/api/v1/articles/?page=1&paginateBy=20&articleNumbers='
#Artikeldaten


#Lizenzartikel Zuordnung 
function getLicensePositions ($articleNr)
{
    $Position = Invoke-RestMethod -Method Get -Uri $URLLicensePosition$articleNr -ContentType 'application/json' -Authentication Bearer -Token $Token -AllowUnencryptedAuthentication
    
    $Position | ForEach-Object {
        getLicensePositonData($Position.data.licenseArticleNumber)
    }
   
}
function getLicensePositonData($licenseArticleNumber) {
    $PositionData = Invoke-RestMethod -Method Get -Uri $URLLicensePositionData$licenseArticleNumber -ContentType 'application/json' -Authentication Bearer -Token $Token -AllowUnencryptedAuthentication
    return $PositionData.data.title
}



