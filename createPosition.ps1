###############################################################
#   Schreibe Auftragsposition
#
#
#
###############################################################
$SQL_Server = 'WWSV05'
$SQL_DB = 'Wiedemann'

function createOrderPosition($i, $OrderNr, $documentType, $customerNumber, $licensingByCustomer) 
{
    $PositionObject = New-Object PSObject
    $PositionObject | Add-Member NoteProperty ARTIKELNR $($ordersComplete.data[0].positions[$i].articleNumber)
    $PositionObject | Add-Member NoteProperty KUNDENNR $customerNumber
    $PositionObject | Add-Member NoteProperty BELEGART $documentType
    $PositionObject | Add-Member NoteProperty BELEGNR $OrderNr
    $PositionObject | Add-Member NoteProperty BEZEICHNUNG $($ordersComplete.data[0].positions[$i].articleTitle)
    $PositionObject | Add-Member NoteProperty CODE1 $($ordersComplete.data[0].positions[$i].mainSupplierNumber)
    $PositionObject | Add-Member NoteProperty POSITIONSNR $(($ordersComplete.data[0].positions[$i].positionNumber) + 1)
    $PositionObject | Add-Member NoteProperty POSTEXT $($ordersComplete.data[0].positions[$i].positionText)
    $PositionObject | Add-Member NoteProperty FIXPOSNR $(($ordersComplete.data[0].positions[$i].fixedPositionNumber) +1 )
    $PositionObject | Add-Member NoteProperty VK $($ordersComplete.data[0].positions[$i].sellingPrice)
    $PositionObject | Add-Member NoteProperty VKPRO $($ordersComplete.data[0].positions[$i].sellingPriceByValue)
    $PositionObject | Add-Member NoteProperty EINHEITVK $($ordersComplete.data[0].positions[$i].sellingPriceByUnit)
    $PositionObject | Add-Member NoteProperty LEK $($ordersComplete.data[0].positions[$i].purchasePrice)
    $PositionObject | Add-Member NoteProperty MENGE_BESTELLT $($ordersComplete.data[0].positions[$i].quantityOrdered)
    $PositionObject | Add-Member NoteProperty MENGE_GELIEFERT $($ordersComplete.data[0].positions[$i].quantityOrdered)
    $PositionObject | Add-Member NoteProperty STATUS 2
    $PositionObject | Add-Member NoteProperty STEUERSCHLUESSEL $(getTaxData $customerNumber $($ordersComplete.data[0].positions[$i].articleGroup) 'MwStKZ')
    $PositionObject | Add-Member NoteProperty TaxRate $(getTaxData $customerNumber $ordersComplete.data[0].positions[$i].articleGroup 'TaxRate')
    $PositionObject | Add-Member NoteProperty RABATTPROZ1 $($ordersComplete.data[0].positions[$i].discountPercent1)
    $PositionObject | Add-Member NoteProperty RABATTPROZ2 $($ordersComplete.data[0].positions[$i].discountPercent2)
    $PositionObject | Add-Member NoteProperty RABATTPROZ3 $($ordersComplete.data[0].positions[$i].discountPercent3)
    $PositionObject | Add-Member NoteProperty POSITIONSART 'P'
    $PositionObject | Add-Member NoteProperty POSBETRAG $($ordersComplete.data[0].positions[$i].positionAmount)
    $PositionObject | Add-Member NoteProperty POSERTRAG $($ordersComplete.data[0].positions[$i].positionProfit)
    $PositionObject | Add-Member NoteProperty WVP $($ordersComplete.data[0].positions[$i].resalePrice)
    $PositionObject | Add-Member NoteProperty C80_Licensing $(isLicenseArticle $customerNumber $($ordersComplete.data[0].positions[$i].articleNumber) $licensingByCustomer)
    $PositionObject | Add-Member NoteProperty C80_RVM $($ordersComplete.data[0].positions[$i].traceabilityFeature)
    $PositionObject | Add-Member NoteProperty FAKTORVK 1.000000000000000
    $PositionObject | Add-Member NoteProperty PREISFAKTOR 1.000000000000000
    $PositionObject | Add-Member NoteProperty ERLOESKONTO $(getTaxData $customerNumber $ordersComplete.data[0].positions[$i].articleGroup 'Account')
    $PositionObject | Add-Member NoteProperty LAGERNR 1
    $PositionObject | Add-Member NoteProperty KW $KW
    $PositionObject | Add-Member NoteProperty ClerkID $(getClerk $VertreterNr )


    

    #Write-Host $($ordersComplete.data[0].positions[$i].articleGroup)

    #Write-Host $PositionObject 


    $SQL_Insert_Colums = @()
    $SQL_Insert_Values = @()

    #SQL table colum name
    $PositionObject | Get-Member -MemberType NoteProperty | Select-Object -Property Name | ForEach-Object { $SQL_Insert_Colums += $_.Name } 

    #SQL table colum values 
    $SQL_Insert_Colums | ForEach-Object { 
        switch($_)
        {
            #'C80_Licensing' { if ($($PositionObject | Select-Object -ExpandProperty $_ ) -eq $False ) { $SQL_Insert_Values += 0} else { $SQL_Insert_Values += 1 }} 
            'FIXPOSNR' { $SQL_Insert_Values += $($PositionObject | Select-Object -ExpandProperty $_) }
            'POSITIONSNR' { $SQL_Insert_Values += $($PositionObject | Select-Object -ExpandProperty $_) }
            'VKPRO' { $SQL_Insert_Values += $($PositionObject | Select-Object -ExpandProperty $_) }
            default {$SQL_Insert_Values += ("'" + $(checkPosOrderValue($($PositionObject | Select-Object -ExpandProperty $_ ))) + "'")}
        }
         
    }



    $SQL_Insert_Colums = "("+($SQL_Insert_Colums -join ", ") + ")"

    #Write-Host $SQL_Insert_Colums

    $SQL_Insert_Values = "("+($SQL_Insert_Values -join ", ") + ")"

    #Write-Host $SQL_Insert_Values

    $SQL_Pos_Insert_Query = "INSERT INTO AUFTRAGSPOS $SQL_Insert_Colums VALUES $SQL_Insert_Values"

    #Write-Host $SQL_Pos_Insert_Query

    Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Pos_Insert_Query
   
}

function checkPosOrderValue($posValue)
{
    if (-not [string]::IsNullOrEmpty($posValue)) { 

        return $posValue 
    }
    else { return [System.DBNull]::Value }
}