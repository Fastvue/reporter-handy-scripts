<# Description:
This script closes all open indexes (dates) in Fastvue Reporter's Elasticsearch database.
Closing indexes frees up memory resources, and Fastvue Reporter will automatically reopen indexes when needed. 
You'll notice today's and yesterday's indexes will automatically re-open after running this script, 
as Fastvue Reporter opens these to import logs and run queries for the live dashboards.

Instructions:
1. In Fastvue Reporter, go to Settings > Diagnostic > Database and note the Cluster URI
2. Enter the cluster URI in the $ElasticsearchUrl variable below
3. Log into the Fastvue server and run this script in PowerShell

#>

# Define the Elasticsearch URL
$ElasticsearchUrl = "http://localhost:9200" # Yours will be different. Get this from Settings > Diagnostic > Database in Fastvue Reporter.

 #### Do not modify below this line ####

 $CloseIndicesEndpoint = "$ElasticsearchUrl/fastvue-records-*/_close"

try {
    # Send the POST request
    $Response = Invoke-RestMethod -Uri $CloseIndicesEndpoint -Method Post -ContentType "application/json"

    # Check the response and display the result
    if ($Response.acknowledged -eq $true) {
        Write-Host "Successfully closed indices." -ForegroundColor Green
    } else {
        Write-Host "Failed to close indices. Response:" -ForegroundColor Red
        $Response | ConvertTo-Json -Depth 10
    }
} catch {
    # Catch and display any errors
    Write-Host "An error occurred:" -ForegroundColor Red
    Write-Host $_.Exception.Message
}
