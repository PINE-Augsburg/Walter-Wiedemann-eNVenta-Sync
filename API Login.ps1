########################################################################
# Walter Wiedemannn 
# 
# API Login mit Bearer Token
########################################################################


#Funktion Includes

. O:\Entwicklung\eNVenta_REST\createPosition.ps1
. O:\Entwicklung\eNVenta_REST\createOrderHead.ps1
. O:\Entwicklung\eNVenta_REST\getNextOrderNumber.ps1
. O:\Entwicklung\eNVenta_REST\getNextOfferNumber.ps1
. O:\Entwicklung\eNVenta_REST\updateOrder.ps1
. O:\Entwicklung\eNVenta_REST\getTaxData.ps1
. O:\Entwicklung\eNVenta_REST\getClerk.ps1
. O:\Entwicklung\eNVenta_REST\isLicenseCustomer.ps1

$SQL_Server = 'WWSV05'
$SQL_DB = 'Wiedemann'

function updateOrder($countedPositions, $orderId, $OrderNr) {
    $URL_Update_Order = "https://api.production.app-project2028.wiedemann-augsburg.de/api/v1/acquisition-sales-orders/$orderId/erp-processed"

    $SQL_Query_Check_Order = "SELECT COUNT(POSITIONSNR) FROM AUFTRAGSPOS WHERE BELEGNR=$OrderNr"
    $SQL_Query_FSROWID = "SELECT FSROWID, ROWID FROM AUFTRAGSKOPF WHERE BELEGNR=$OrderNr"

    $countedSQLPositions = Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Query_Check_Order
    $orderHeadRowIds = Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Query_FSROWID

    $updateHeader = @{
        "erpDocumentNumber" = $OrderNr
        "erpFsRowId" = $orderHeadRowIds.Item(0)
        "erpRowId" = $orderHeadRowIds.Item(1)
    } 
    
    $updateHeaderJSON = $updateHeader | ConvertTo-Json

    if ($countedPositions -eq $countedSQLPositions.Item(0))
    {
        Write-Host "Auftrag wurde erfolgreich angelegt!" -ForegroundColor Green
        Invoke-RestMethod $URL_Update_Order -ContentType 'application/json' -Method Post -Body $updateHeaderJSON -Authentication Bearer -Token $Token -AllowUnencryptedAuthentication

    }
    else 
    {
        Write-Host "Auftrag mit ID $orderId konnte nicht angelegt werden!" -ForegroundColor Red
    }

    #Send-MailMessage -From 'app-project2028@wiedemann-augsburg.de' -To 'markus.probst@wiedemann-augsburg.de' -Subject 'I am alive - Greetings from your sync!' -SmtpServer 'wiedemannaugsburg-de02c.mail.protection.outlook.com'

    #Debug
    #----------------------------------------------------
    #Write-Host $updateHeader.erpDocumentNumber
    #Write-Host $URL_Update_Order
    #Invoke-RestMethod $URL_Update_Order -ContentType 'application/json' -Method Post -Body $updateHeaderJSON -Authentication Bearer -Token $Token -AllowUnencryptedAuthentication
}

#URL Login
$URLLogin = 'https://api.production.app-project2028.wiedemann-augsburg.de/api/v1/auth/login/'

#URL Neue AuftrÃ¤ge mit Status 'released'
$URLNewOrders = 'https://api.production.app-project2028.wiedemann-augsburg.de/api/v1/acquisition-sales-orders/?page=1&paginateBy=20&stati=released'

#Benutzer Credentials
$header = @{
    username='admin'
    password='Oow.eighash.ah7dieM0'
}



#Login = TRUE -> Response Token 
$secureToken = Invoke-RestMethod $URLLogin -ContentType 'application/x-www-form-urlencoded' -Method Post -Body $header

#Deklarierung in SecureString 
$Token = $secureToken.access_token | ConvertTo-SecureString -AsPlainText -Force

$ordersComplete = Invoke-RestMethod -Method Get -Uri $URLNewOrders -ContentType 'application/json' -Authentication Bearer -Token $Token -AllowUnencryptedAuthentication


$orderCheckTime = Get-Date

if ($ordersComplete.data.Length -gt 0)
{

    Write-Host '========================================================================' -ForegroundColor Green 
    Write-Host '[WaWi-Sync] Es stehen'  $ordersComplete.data.Length 'Aufträge zur Synchronisierung aus!' -ForegroundColor Green 
    Write-Host '[WaWi-Sync]' $orderCheckTime -ForegroundColor Green 
    Write-Host '========================================================================' -ForegroundColor Green 


        $VertreterNr = $ordersComplete.data[0].headData.salesmanNumber1
        $documentType = $ordersComplete.data[0].headData.documentType
        $Kunde = $ordersComplete.data[0].headData.invoiceCompanyName1 
        $customerNumber = $ordersComplete.data[0].headData.customerNumber
        $createdAt = Get-Date($ordersComplete.data[0].createdAt) -Format "dd/MM/yyyy HH:mm:ss"
        $modifiedAt = Get-Date($ordersComplete.data[0].modifiedAt) -Format "dd/MM/yyyy HH:mm:ss"
        $KW = Get-Date($ordersComplete.data[0].createdAt) -UFormat %V
        
        if($documentType -eq 1)
        {
            $newOrderNumber = getOrderNumber($VertreterNr)
        }
        else{
            $newOrderNumber = getOfferNumber($VertreterNr)
        }
        
        $orderID = $ordersComplete.data[0]._id
        $clerkID = $ordersComplete.data[0].actor.clerkId
        $countedPositions = $ordersComplete.data[0].positions.length
        #$vatID = $odersComplete.data[0].headData.vatId

        Write-Host ""
        Write-Host "Kunde:" $Kunde -ForegroundColor Yellow
        Write-Host "Vertreter:" $VertreterNr $clerkID $(getClerk($VertreterNr)) -ForegroundColor Yellow
        Write-Host "Auftragsart: " $(if ($documentType -eq 1) {"Auftrag"} else {"Angebot"}) -ForegroundColor Yellow
        Write-Host "Anzahl Positionen:" $countedPositions -ForegroundColor Yellow
        Write-Host "Generierte Auftrags-Nr:" $newOrderNumber -ForegroundColor Yellow
        Write-Host "Auftrag vom: " $createdAt -ForegroundColor Yellow
        Write-Host "Auftrags-ID: " $orderID -ForegroundColor Yellow
        Write-Host "" 
        #Überprüfung Lizenzierungsauftrag

        $licensingByCustomer = isLicenseCustomer $ordersComplete.data[0].headData.customerNumber

        #Verabeite Auftragskopf
        createOrderHead 0 $newOrderNumber $createdAt $modifiedAt $KW $VertreterNr $licensingByCustomer

        
        for ($i=0; $i -lt $ordersComplete.data[0].positions.length; $i++)
        {
            #ErgÃ¤nze Beleg-Nr und schreibe Auftragspositionen
            createOrderPosition $i $newOrderNumber $documentType $customerNumber $createdAt $modifiedAt $KW $VertreterNr $licensingByCustomer
        }

        #Check Order Positionen in SQL Datenbank und aktualisiere Status in API 
        updateOrder $countedPositions $orderId $newOrderNumber
}
else 
{

    Write-Host '========================================================================' -ForegroundColor Green 
    Write-Host '[WaWi-Sync] Es gibt keine neuen AuftrÃ¤ge!' -ForegroundColor Green 
    Write-Host '[WaWi-Sync]' $orderCheckTime -ForegroundColor Green 
    Write-Host '========================================================================' -ForegroundColor Green 
}



function updateOrder($countedPositions, $orderId, $OrderNr) {
    $URL_Update_Order = "https://api.production.app-project2028.wiedemann-augsburg.de/api/v1/acquisition-sales-orders/$orderId/erp-processed"

    $SQL_Query_Check_Order = "SELECT COUNT(POSITIONSNR) FROM AUFTRAGSPOS WHERE BELEGNR=$OrderNr"
    $SQL_Query_FSROWID = "SELECT FSROWID, ROWID FROM AUFTRAGSKOPF WHERE BELEGNR=$OrderNr"

    $countedSQLPositions = Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Query_Check_Order
    $orderHeadRowIds = Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Query_FSROWID

    $updateHeader = @{
        "erpDocumentNumber" = $OrderNr
        "erpFsRowId" = $orderHeadRowIds.Item(0)
        "erpRowId" = $orderHeadRowIds.Item(1)
    } 
    
    $updateHeaderJSON = $updateHeader | ConvertTo-Json

    if ($countedPositions -eq $countedSQLPositions.Item(0))
    {
        Write-Host "Auftrag wurde erfolgreich angelegt!" -ForegroundColor Green
        Invoke-RestMethod $URL_Update_Order -ContentType 'application/json' -Method Post -Body $updateHeaderJSON -Authentication Bearer -Token $Token -AllowUnencryptedAuthentication

    }
    else 
    {
        Write-Host "Auftrag mit ID $orderId konnte nicht angelegt werden!" -ForegroundColor Red
    }

    #Send-MailMessage -From 'app-project2028@wiedemann-augsburg.de' -To 'markus.probst@wiedemann-augsburg.de' -Subject 'I am alive - Greetings from your sync!' -SmtpServer 'wiedemannaugsburg-de02c.mail.protection.outlook.com'

    #Debug
    #----------------------------------------------------
    #Write-Host $updateHeader.erpDocumentNumber
    #Write-Host $URL_Update_Order
    #Invoke-RestMethod $URL_Update_Order -ContentType 'application/json' -Method Post -Body $updateHeaderJSON -Authentication Bearer -Token $Token -AllowUnencryptedAuthentication
}
