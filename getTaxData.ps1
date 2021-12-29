$SQL_Server = 'WWSV05'
$SQL_DB = 'Wiedemann'



function getTaxData($customerNumber, $ArticleGroup, $mode) {
   
    #MWSTKZ von Kundenstamm
    $SQL_Query_MwstKZ = "SELECT MWSTKZ FROM KUNDEN WHERE KUNDENNR=$customerNumber"
    $MwstKZ = Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Query_MwstKZ
    $vatID = $MwstKZ.Item(0)
    $SQL_Query_TaxData = "SELECT TaxCodeSales, TaxCodePurchase, AccountSales FROM VATArticleStorage WHERE VATID=$vatID AND ArticleGroup=$ArticleGroup"
    #Write-Host $SQL_Query_TaxData
    $TaxData = Invoke-Sqlcmd -ServerInstance $SQL_Server -Database $SQL_DB -Query $SQL_Query_TaxData


    #Write-Host $TaxData.Item("TaxCodeSales")
    switch ($TaxData.Item("TaxCodePurchase"))
    {
        "V19" {
            $TaxData.Item("TaxCodePurchase") = '19.00'
            break;
        }
        "V07" {
            $TaxData.Item("TaxCodePurchase") = '7.00'
            break;
        }
        "V00" {
            $TaxData.Item("TaxCodePurchase") = '0.00'
            break; 
        }
    }
    #Write-Host $TaxData.Item("TaxCodePurchase")
    #Write-Host $TaxData.Item("AccountSales")

    switch($mode) {
        'MwStKZ' {
            return $TaxData.Item("TaxCodeSales")
            break;
        }
        'Account' {
            return $TaxData.Item("AccountSales")
            break;
        }
        'TaxRate' {
            return $TaxData.Item("TaxCodePurchase")
            break;
        }

    }
}


#getTaxData 23652 20 'TaxRate'

