Bulk PageSpeed Insights Automation for Google Sheets

This Google Apps Script automates the process of checking URL status codes and fetching PageSpeed Insights (PSI) performance metrics for multiple URLs directly from Google Sheets.

Features

✅ Validate URLs by checking HTTP status codes (200, 301 marked as valid)
✅ Separate sheets for Mobile and Desktop PSI results
✅ Automatically creates headers for performance metrics:

Score

First Contentful Paint

Speed Index

Largest Contentful Paint

Time to Interactive

Total Blocking Time

Cumulative Layout Shift
✅ Marks invalid URLs in red
✅ Uses Google PageSpeed Insights API

Prerequisites

A Google Sheet with:

Sheet Name: InputSheet

Column A: List of URLs (starting from A2)

A Google PageSpeed Insights API Key
Get it from: Google Cloud Console

Setup Instructions

Open Google Sheets.

Go to Extensions → Apps Script.

Paste the script into the Apps Script editor.

Replace the value of apikey in the PageSpeedApiEndpointUrl function with your Google API key.

Save the script.

How to Use

In the InputSheet, list all URLs starting from cell A2.

Run the function GeneratePages() from the Apps Script editor.

The script will:

Check status codes for all URLs and update Column B with:

✅ Black text for valid URLs (status 200 or 301)

❌ Red text for invalid URLs

Create two sheets:

Mobile → PSI metrics for mobile strategy

Desktop → PSI metrics for desktop strategy

PSI results for valid URLs will populate automatically in respective sheets.

Functions Overview

GeneratePages()
Main function that:

Reads URLs

Checks HTTP status codes

Creates Mobile & Desktop sheets

Fetches PSI metrics and populates sheets

PageSpeedApiEndpointUrl(strategy, url)
Builds the Google PSI API endpoint URL for a given strategy (mobile or desktop).

getStatusCode(url)
Returns HTTP status code for the given URL.

fetchDataFromPSI(strategy, url)
Calls the PSI API and extracts the following metrics:

Performance Score

First Contentful Paint

Speed Index

Largest Contentful Paint

Time to Interactive

Total Blocking Time

Cumulative Layout Shift

Example Sheet Structure
InputSheet
URL	Status
https://example.com
	200
https://invalid.com
	404
Mobile / Desktop
URL	Score	FCP	SI	LCP	TTI	TBT	CLS
https://example.com
	85	1.2 s	3.0 s	2.1 s	2.8 s	150 ms	0.05
