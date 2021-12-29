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
$orderNr_Boeck_Start = "2" + $orderNrYearIndex + "47000"
$orderNr_Boeck_End = "2" + $orderNrYearIndex + "53999"

$orderNr_Kotter_Start = "2" + $orderNrYearIndex + "54000"
$orderNr_Kotter_End = "2" + $orderNrYearIndex + "59999"

$orderNr_Quartal_Start = "2" + $orderNrYearIndex + "60000"
$orderNr_Quartal_End = "2" + $orderNrYearIndex + "64999"

$orderNr_Keller_Start = "2" + $orderNrYearIndex + "65000"
$orderNr_Keller_End = "2" + $orderNrYearIndex + "69999"

$orderNr_Rinkenburger_Start = "2" + $orderNrYearIndex + "70000" 
$orderNr_Rinkenburger_End = "2" + $orderNrYearIndex + "74999" 

$orderNr_Brueggemann_Start = "2" + $orderNrYearIndex + "81000" 
$orderNr_Brueggemann_End = "2" + $orderNrYearIndex + "84999" 

<#
Write-Host "BOECK: $orderNr_Boeck_Start "-" $orderNr_Boeck_End" 
Write-Host "KOTTER: $orderNr_Kotter_Start "-" $orderNr_Kotter_End"
Write-Host "QUARTAL: $orderNr_Quartal_Start "-" $orderNr_Quartal_End"
Write-Host "RINKENBURGER: $orderNr_Rinkenburger_Start "-" $orderNr_Rinkenburger_End"
Write-Host "BRUEGGEMANN": $orderNr_Brueggemann_Start "-" $orderNr_Brueggemann_End
#>

function getOrderNumber($SalesClerk) {
    
    $nextOrderNr = ""

    switch ($SalesClerk)
    {
        80 {
             
             $nextOrderNr =  Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $orderNr_Boeck_Start 'AND' $orderNr_Boeck_End ' ORDER BY BELEGNR DESC"
             return $nextOrderNr.BELEGNR + 1
        }
        36 {
             
              $nextOrderNr = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $orderNr_Kotter_Start 'AND' $orderNr_Kotter_End' ORDER BY BELEGNR DESC"
              return $nextOrderNr.BELEGNR + 1
        }
        55 {

              $nextOrderNr = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $orderNr_Quartal_Start 'AND' $orderNr_Quartal_End' ORDER BY BELEGNR DESC"
              return $nextOrderNr.BELEGNR + 1
        }
        70 {

              $nextOrderNr = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $orderNr_Keller_Start 'AND' $orderNr_Keller_End' ORDER BY BELEGNR DESC"
              return $nextOrderNr.BELEGNR + 1
        }
        45 {

              $nextOrderNr = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $orderNr_Rinkenburger_Start 'AND' $orderNr_Rinkenburger_End' ORDER BY BELEGNR DESC"
              return $nextOrderNr.BELEGNR + 1
        }
        90 {

              $nextOrderNr = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $SqlDB -Query "SELECT TOP 1 BELEGNR FROM AUFTRAGSKOPF WHERE BELEGNR BETWEEN ' $orderNr_Brueggemann_Start 'AND' $orderNr_Brueggemann_End' ORDER BY BELEGNR DESC"
              return $nextOrderNr.BELEGNR + 1
        }
        Default 
        {"Error"}
    }
}

