###############################################################
#   Beziehe nächst frei Auftragsnummer
#
#
#
###############################################################


#Jahreszahl
$orderNrYearIndex = Get-Date -Format yy

#SQL Server Verbindung 
$SqlServer = "WWSV05"
$SqlDB = 'Wiedemann'

#Nummernkreise 
$offerNr_Boeck_Start = "2" + $orderNrYearIndex + "94700"
$offerNr_Boeck_End = "2" + $orderNrYearIndex + "95399"

$offerNr_Kotter_Start = "2" + $orderNrYearIndex + "95400"
$offerNr_Kotter_End = "2" + $orderNrYearIndex + "95999"

$offerNr_Quartal_Start = "2" + $orderNrYearIndex + "96000"
$offerNr_Quartal_End = "2" + $orderNrYearIndex + "96499"

$offerNr_Keller_Start = "2" + $orderNrYearIndex + "96500"
$offerNr_Keller_End = "2" + $orderNrYearIndex + "96999"

$offerNr_Rinkenburger_Start = "2" + $orderNrYearIndex + "97000" 
$offerNr_Rinkenburger_End = "2" + $orderNrYearIndex + "97499" 

$offerNr_Brueggemann_Start = "2" + $orderNrYearIndex + "98100" 
$offerNr_Brueggemann_End = "2" + $orderNrYearIndex + "98499" 

<#
Write-Host "BOECK: $orderNr_Boeck_Start "-" $orderNr_Boeck_End" 
Write-Host "KOTTER: $orderNr_Kotter_Start "-" $orderNr_Kotter_End"
Write-Host "QUARTAL: $orderNr_Quartal_Start "-" $orderNr_Quartal_End"
Write-Host "RINKENBURGER: $orderNr_Rinkenburger_Start "-" $orderNr_Rinkenburger_End"
Write-Host "BRUEGGEMANN": $orderNr_Brueggemann_Start "-" $orderNr_Brueggemann_End
#>

function getOfferNumber($SalesClerk) {
    
    $nextOfferNr = ""

    switch ($SalesClerk)
    {
        80 {
             
            $nextOfferNr =  Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $offerNr_Boeck_Start 'AND' $offerNr_Boeck_End ' ORDER BY BELEGNR DESC"
            
            return $nextOfferNr.BELEGNR + 1
        }
        36 {
             
            $nextOfferNr = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $offerNr_Kotter_Start 'AND' $offerNr_Kotter_End' ORDER BY BELEGNR DESC"
            return $nextOfferNr.BELEGNR + 1
        }
        55 {

            $nextOfferNr = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $offerNr_Quartal_Start 'AND' $offerNr_Quartal_End' ORDER BY BELEGNR DESC"
            return $nextOfferNr.BELEGNR + 1
        }
        70 {

            $nextOfferNr = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $offerNr_Keller_Start 'AND' $offerNr_Keller_End' ORDER BY BELEGNR DESC"
            return $nextOfferNr.BELEGNR + 1
        }
        45 {

            $nextOfferNr = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $offerNr_Rinkenburger_Start 'AND' $offerNr_Rinkenburger_End' ORDER BY BELEGNR DESC"
            return $nextOfferNr.BELEGNR + 1
        }
        90 {

            $nextOfferNr = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $offerNr_Brueggemann_Start 'AND' $offerNr_Brueggemann_End' ORDER BY BELEGNR DESC"
            return $nextOfferNr.BELEGNR + 1
        }
        Default 
        {"Error"}
    }
}
