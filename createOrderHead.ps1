###############################################################
#   Schreibe Auftragsposition
#
#
#
###############################################################
$SQL_Server = 'WWSV05'
$SQL_DB = 'Wiedemann'
function createOrderHead($i, $newOrderNumber, $createdAt, $modifiedAt, $KW, $VertreterNr, $licensingByCustomer) 
{
    $OrderHeadObj = New-Object PSObject
    $OrderHeadObj | Add-Member Noteproperty BELEGNR $newOrderNumber 
    $OrderHeadObj | Add-Member Noteproperty BELEGART $($ordersComplete.data[$i].headData.documentType)
    $OrderHeadObj | Add-Member Noteproperty BESTELLART $($ordersComplete.data[$i].headData.orderType)
    $OrderHeadObj | Add-Member Noteproperty KUNDENNR $($ordersComplete.data[$i].headData.customerNumber)
    $OrderHeadObj | Add-Member Noteproperty VATID $($ordersComplete.data[$i].headData.vatId)
    $OrderHeadObj | Add-Member Noteproperty VATRegNo $($ordersComplete.data[$i].headData.vatRegistrationNumber)
    $OrderHeadObj | Add-Member Noteproperty RFIRMA1 $($ordersComplete.data[$i].headData.invoiceCompanyName1)
    $OrderHeadObj | Add-Member Noteproperty RFIRMA2 $($ordersComplete.data[$i].headData.invoiceCompanyName2)
    $OrderHeadObj | Add-Member Noteproperty RSTRASSE $($ordersComplete.data[$i].headData.invoiceAddress)
    $OrderHeadObj | Add-Member Noteproperty RPLZ $($ordersComplete.data[$i].headData.invoicePostalCode)
    $OrderHeadObj | Add-Member Noteproperty RORT $($ordersComplete.data[$i].headData.invoiceCity)
    $OrderHeadObj | Add-Member Noteproperty RLAENDERKZ $($ordersComplete.data[$i].headData.invoiceCountryCode)
    $OrderHeadObj | Add-Member Noteproperty RLAND $($ordersComplete.data[$i].headData.invoiceCountry)
    $OrderHeadObj | Add-Member Noteproperty LFIRMA1 $($ordersComplete.data[$i].headData.deliveryCompanyName1)
    $OrderHeadObj | Add-Member Noteproperty LFIRMA2 $($ordersComplete.data[$i].headData.deliveryCompanyName2)
    $OrderHeadObj | Add-Member Noteproperty LSTRASSE $($ordersComplete.data[$i].headData.deliveryAddress)
    $OrderHeadObj | Add-Member Noteproperty LORT $($ordersComplete.data[$i].headData.deliveryCity)
    $OrderHeadObj | Add-Member Noteproperty LPLZ $($ordersComplete.data[$i].headData.deliveryPostalCode)
    $OrderHeadObj | Add-Member Noteproperty LLAENDERKZ $($ordersComplete.data[$i].headData.deliveryCountryCode)
    $OrderHeadObj | Add-Member Noteproperty LLAND $($ordersComplete.data[$i].headData.deliveryCountry)
    $OrderHeadObj | Add-Member Noteproperty ANSCHRIFTNR $($ordersComplete.data[$i].headData.deliveryAddressNumber)
    $OrderHeadObj | Add-Member Noteproperty SPRACHE $($ordersComplete.data[$i].headData.customerLanguage)
    $OrderHeadObj | Add-Member Noteproperty STATUS 1
    $OrderHeadObj | Add-Member Noteproperty SPERRUNG 0
    $OrderHeadObj | Add-Member Noteproperty WAEHRUNG $($ordersComplete.data[$i].headData.currency)
    $OrderHeadObj | Add-Member Noteproperty MEMO $($ordersComplete.data[$i].headData.memo)
    $OrderHeadObj | Add-Member Noteproperty STARTTEXT $($ordersComplete.data[$i].headData.startText)
    $OrderHeadObj | Add-Member Noteproperty ENDETEXT $($ordersComplete.data[$i].headData.endText)
    $OrderHeadObj | Add-Member Noteproperty VERTRETERNR1 $($ordersComplete.data[$i].headData.salesmanNumber1)
    $OrderHeadObj | Add-Member Noteproperty SAMMELRECHNUNG $($ordersComplete.data[$i].headData.useCollectiveInvoice)
    $OrderHeadObj | Add-Member Noteproperty StartTextIsHTML $($ordersComplete.data[$i].headData.startTextIsHtml)
    $OrderHeadObj | Add-Member Noteproperty EndTextIsHTML $($ordersComplete.data[$i].headData.endTextIsHtml)
    $OrderHeadObj | Add-Member Noteproperty C80_Licensing $licensingByCustomer
    $OrderHeadObj | Add-Member Noteproperty C80_CollectiveDeliveryNote $($ordersComplete.data[$i].headData.useCollectiveDeliveryNote)
    $OrderHeadObj | Add-Member Noteproperty C80_CollectivePackingList $($ordersComplete.data[$i].headData.useCollectivePackingList)
    $OrderHeadObj | Add-Member Noteproperty ERFASSUNGSDATUM $createdAt
    $OrderHeadObj | Add-Member Noteproperty AENDERUNGSDATUM $modifiedAt
    $OrderHeadObj | Add-Member Noteproperty ReleaseDate $createdAt
    $OrderHeadObj | Add-Member Noteproperty BESTELLDATUM $createdAt
    $OrderHeadObj | Add-Member Noteproperty NV4_CreatedByUser $($ordersComplete.data[$i].actor.clerkId)
    $OrderHeadObj | Add-Member Noteproperty ClerkIDEntry $($ordersComplete.data[$i].actor.clerkId)
    $OrderHeadObj | Add-Member Noteproperty ReleaseClerk $($ordersComplete.data[$i].actor.clerkId)
    $OrderHeadObj | Add-Member Noteproperty SACHBEARBEITERNR $(getClerk($VertreterNr))
    $OrderHeadObj | Add-Member Noteproperty KURS 1.0000000
    $OrderHeadObj | Add-Member NoteProperty ZBNUMMER $($ordersComplete.data[$i].headData.paymentTermNumber)
    $OrderHeadObj | Add-Member NoteProperty LBNUMMER $($ordersComplete.data[$i].headData.deliveryTermNumber)
    $OrderHeadObj | Add-Member NoteProperty KW $KW
   
    #Write-Host $OrderHeadObj

    $SQL_Head_Colums = @()
    $SQL_Head_Colums_Values = @()

    #SQL Table Colum Name
    $OrderHeadObj | Get-Member -MemberType NoteProperty | Select-Object -Property Name | ForEach-Object { $SQL_Head_Colums += $_.Name}

    #Write-Host $OrderHeadObj
    #SQL Table Colum Values

    $SQL_Head_Colums | ForEach-Object { 
        switch ($_)
        {
            'C80_CollectiveDeliveryNote' { if ($($OrderHeadObj | Select-Object -ExpandProperty $_ ) -eq $False ) { $SQL_Head_Colums_Values += 0} else { $SQL_Head_Colums_Values += 1 }}
            'C80_CollectivePackingList' { if ($($OrderHeadObj| Select-Object -ExpandProperty $_ ) -eq $False ) { $SQL_Head_Colums_Values += 0 } else { $SQL_Head_Colums_Values += 1 }}
            'StartTextIsHTML' { if ($($OrderHeadObj | Select-Object -ExpandProperty $_ ) -eq $True ) { $SQL_Head_Colums_Values += 1} else { $SQL_Head_Colums_Values += 0 }}
            'EndTextIsHTML' { if ($($OrderHeadObj | Select-Object -ExpandProperty $_ ) -eq $True ) { $SQL_Head_Colums_Values += 1} else { $SQL_Head_Colums_Values += 0 }}
            'SAMMELRECHNUNG' { if ($($OrderHeadObj | Select-Object -ExpandProperty $_ ) -eq $True ) { $SQL_Head_Colums_Values += 1} else { $SQL_Head_Colums_Values += 0 }}
            default { $SQL_Head_Colums_Values += ("'" + $(checkHeadOrderValue($($OrderHeadObj | Select-Object -ExpandProperty $_))) + "'")}
        }
    }

    $SQL_Head_Colums = "("+($SQL_Head_Colums -join ", ") + ")"
    $SQL_Head_Colums_Values = "("+($SQL_Head_Colums_Values -join ", ") + ")"
    #Write-Host $SQL_Head_Colums
    #Write-Host $SQL_Head_Colums_Values 

    $SQL_Head_Insert_Query = "INSERT INTO AUFTRAGSKOPF $SQL_Head_Colums VALUES $SQL_Head_Colums_Values"

    #Write-Host $SQL_Head_Insert_Query
    Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Head_Insert_Query

}

function checkHeadOrderValue($headValue)
{
    if (-not [string]::IsNullOrEmpty($headValue)) { 

        return $headValue 
    }
    else { return [System.DBNull]::Value }
}
