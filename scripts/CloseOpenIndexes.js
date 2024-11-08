;(async function closeOpenIndexes() {
	/*
  Description:
  This script closes all open indexes (older dates) in Fastvue Reporter's Elasticsearch database, excluding today’s and yesterday’s indexes.
  Closing indexes frees up memory resources, and Fastvue Reporter will automatically reopen these indexes when needed (e.g., when running reports on older dates).

  Instructions:
  1. Update the `baseURL` variable to match your Fastvue Reporter URL.
  2. In Chrome, go to your Fastvue Reporter's web interface, and navigate to **Settings > Diagnostics > Database**.
     - If the database status is "Bad" or "Unknown," restart the Fastvue Reporter service in **services.msc**
     - Wait for the status to show either "Connected", "Waiting for index recovery", "Preparing Elasticsearch" or 'Operational"
  3. Open Chrome DevTools:
     - Press `F12` or `Ctrl+Shift+I` (Windows/Linux) or `Cmd+Option+I` (Mac).
  4. Go to the **Console** tab, paste the script, and press `Enter`.
*/

	// Update this base URL to match your Fastvue Reporter site
	const baseURL = "http://your_fastvue_site"

	// ---- Do not modify below this line (unless you know what you're doing!) ------------------------------

	// Helper function to get the formatted date string for today and yesterday
	function getFormattedDate(daysAgo = 0) {
		const date = new Date()
		date.setDate(date.getDate() - daysAgo)
		return date.toISOString().slice(0, 10).replace(/-/g, "") // Format as YYYYMMDD
	}

	// Dates to exclude (today and yesterday)
	const today = getFormattedDate()
	const yesterday = getFormattedDate(1)

	// Fetch the index information from the API
	const response = await fetch(`${baseURL}/_/api?f=Database.ElasticIndexInfo`)
	const data = await response.json()

	// Check for successful data retrieval
	if (data.Status !== 0) {
		console.error("Failed to retrieve index data")
		return
	}

	// Loop through each index
	for (const index of data.Data) {
		// Check if the index is open and not from today or yesterday
		if (
			index.Status === "open" &&
			index.Name !== `fastvue-records-${today}` &&
			index.Name !== `fastvue-records-${yesterday}`
		) {
			const indexName = index.Name

			// Announce the index about to be closed
			console.log(`Attempting to close index: ${indexName}`)

			try {
				// Send POST request to close the index
				const closeResponse = await fetch(`${baseURL}/_/api?f=Database.ElasticCloseIndex`, {
					method: "POST",
					headers: {
						"Content-Type": "application/json",
					},
					body: JSON.stringify({ index: indexName }),
				})

				const closeData = await closeResponse.json()

				// Check if the close request was successful
				if (closeData.Status === 0 && closeData.Data === true) {
					console.log(`Successfully closed index: ${indexName}`)
				} else {
					console.error(`Failed to close index: ${indexName}`, closeData)
				}
			} catch (error) {
				console.error(`Error closing index: ${indexName}`, error)
			}
		} else if (index.Status === "open") {
			console.log(`Skipped closing index: ${index.Name} (today's or yesterday's index)`)
		}
	}
})()
