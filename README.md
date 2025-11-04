---

# Test Your Website Speed in One Click (Using Sheets)

This Google Apps Script automates the process of checking URL status codes and fetching **PageSpeed Insights (PSI)** performance metrics for multiple URLs directly from **Google Sheets**.

---

## Features

* Validate URLs by checking HTTP status codes (`200`, `301` marked as valid)
* Separate sheets for **Mobile** and **Desktop** PSI results
* Automatically creates headers for performance metrics:

  * **Score**
  * **First Contentful Paint (FCP)**
  * **Speed Index (SI)**
  * **Largest Contentful Paint (LCP)**
  * **Time to Interactive (TTI)**
  * **Total Blocking Time (TBT)**
  * **Cumulative Layout Shift (CLS)**
* Marks invalid URLs in **red**
* Uses **Google PageSpeed Insights API**

---

## Prerequisites

A Google Sheet with:

| Sheet Name     | Description                                                |
| -------------- | ---------------------------------------------------------- |
| **InputSheet** | Contains list of URLs in **Column A** (starting from `A2`) |

You’ll also need a **Google PageSpeed Insights API Key**.
You can get one from the **Google Cloud Console**.

---

## Setup Instructions

1. Open **Google Sheets**.
2. Go to **Extensions → Apps Script**.
3. Paste the script into the **Apps Script editor**.
4. Replace the value of `apikey` in the `PageSpeedApiEndpointUrl()` function with your **Google API key**.
5. Save the script.

---

## How to Use

1. In the **InputSheet**, list all URLs starting from cell `A2`.
2. Run the function `GeneratePages()` from the **Apps Script editor**.

The script will:

* Check status codes for all URLs and update **Column B** with:

  * **Black text** → valid URLs (`200` or `301`)
  * **Red text** → invalid URLs
* Create two sheets:

  * **Mobile** → PSI metrics for mobile strategy
  * **Desktop** → PSI metrics for desktop strategy
* Populate PSI results automatically in respective sheets for valid URLs.

---

## Functions Overview

### `GeneratePages()`

Main function that:

* Reads URLs
* Checks HTTP status codes
* Creates Mobile & Desktop sheets
* Fetches PSI metrics and populates sheets

---

### `PageSpeedApiEndpointUrl(strategy, url)`

Builds the Google PSI API endpoint URL for a given **strategy** (`mobile` or `desktop`).

---

### `getStatusCode(url)`

Returns the **HTTP status code** for the given URL.

---

### `fetchDataFromPSI(strategy, url)`

Calls the PSI API and extracts the following metrics:

| Metric                             | Description                             |
| ---------------------------------- | --------------------------------------- |
| **Performance Score**              | Overall PSI score                       |
| **First Contentful Paint (FCP)**   | Time until first content is painted     |
| **Speed Index (SI)**               | Visual completeness speed               |
| **Largest Contentful Paint (LCP)** | Time until main content is visible      |
| **Time to Interactive (TTI)**      | When the page becomes fully interactive |
| **Total Blocking Time (TBT)**      | Time blocked by long tasks              |
| **Cumulative Layout Shift (CLS)**  | Layout stability                        |

---

## Example Sheet Structure

### InputSheet

| URL                                        | Status |
| ------------------------------------------ | ------ |
| [https://example.com](https://example.com) | 200    |
| [https://invalid.com](https://invalid.com) | 404    |

---

### Mobile / Desktop Sheets

| URL                                        | Score | FCP   | SI    | LCP   | TTI   | TBT    | CLS  |
| ------------------------------------------ | ----- | ----- | ----- | ----- | ----- | ------ | ---- |
| [https://example.com](https://example.com) | 85    | 1.2 s | 3.0 s | 2.1 s | 2.8 s | 150 ms | 0.05 |

---
