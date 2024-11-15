;(async function closeAllOpenIndexes() {
	/*
    Description:
    This script sends a POST request to close all indices matching the pattern "fastvue-records-*"
    in Elasticsearch using the current window's URL as the base.

    Instructions:
		1. In Fastvue Reporter, go to Settings > Diagnostic > Database and note the Cluster URI
    2. Remote onto the Fastvue Server, open Chrome and enter the Cluster URI into the address bar. 
		3. Open Chrome's developer tools by pressing F12 or right-clicking and selecting "Inspect"
		4. Navigate to the "Console" tab
		5. Copy and paste the entire script into the console and press Enter
		6. A success message will be displayed in the console if the indices were closed successfully.
			An error message will be displayed if there was an issue closing the indices.
  */

	try {
		// Construct the URL dynamically using the current window's location
		const baseURL = window.location.origin

		// Send POST request to close all matching indices
		const response = await fetch(`${baseURL}/fastvue-records-*/_close`, {
			method: "POST",
			headers: {
				"Content-Type": "application/json",
			},
		})

		// Parse response
		const data = await response.json()

		// Log result
		if (response.ok) {
			console.log("Successfully closed indices:", data)
		} else {
			console.error("Failed to close indices:", data)
		}
	} catch (error) {
		console.error("Error closing indices:", error)
	}
})()
