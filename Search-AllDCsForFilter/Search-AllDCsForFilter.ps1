$filter = "(proxyAddresses=SMTP:foo@contoso.com)"

$gcs = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().GlobalCatalogs
$gcs | % {
    $searcher = $_.GetDirectorySearcher()
    $searcher.Filter = $filter
    $results = $searcher.FindAll()
    if ($results.Count -gt 0)
    {
        $results
    } 
    else 
    {
        "Not found on GC: " + $_.Name
    }
}

$domains = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Domains
$domains | % {
    $dcs = $_.DomainControllers
    $dcs | % {
        $searcher = $_.GetDirectorySearcher()
        $searcher.Filter = $filter
        $results = $searcher.FindAll()
        if ($results.Count -gt 0)
        {
            $results
        }
        else 
        {
            "Not found on DC: " + $_.Name
        }
    }
}