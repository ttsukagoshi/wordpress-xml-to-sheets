# WordPress XML to Google Sheets
[![GitHub Super-Linter](https://github.com/ttsukagoshi/xml-to-google-sheets/workflows/Lint%20Code%20Base/badge.svg)](https://github.com/marketplace/actions/super-linter) [![Total alerts](https://img.shields.io/lgtm/alerts/g/ttsukagoshi/xml-to-google-sheets.svg?logo=lgtm&logoWidth=18)](https://lgtm.com/projects/g/ttsukagoshi/xml-to-google-sheets/alerts/)

Convert XML exported from WordPress onto Google Sheets.

## About
A Spreadsheet-bound apps script solution to convert XML exported from the native WordPress tool into your Google Sheets file.

<blockquote>
<p>This is a solution with a very limited target: if you want to export the contents of your WordPress-powered website into your spreadsheet, using existing WordPress plugins should be your first choice. Crawl the web for keywords like <code>wordpress</code> and <code>csv</code>. Or if you are using WordPress versions 4.7 or later, you should be able to take advantage of the <a href="https://developer.wordpress.org/rest-api/">WordPress REST API</a>, which, as part of the WordPress core features, would provide you with a more integrated experience.</p>
<p>This particular solution can be useful for those who do not/cannot install new plugins for some reason, and/or those who are using old versions of WordPress that are too old for newer features.</p>
</blockquote>

## Install
You can either
1. Copy [this sample spreadsheet](https://docs.google.com/spreadsheets/d/1n8KSY_wBYDxW-6pTrS3FwOpRrvcv9Q48o9H488C3upE/edit#gid=1746753348) or 
2. Copy & paste the file contents inside [the `src` directory of the GitHub repository](https://github.com/ttsukagoshi/wordpress-xml-to-sheets/tree/main/src) using [Google Sheets' script editor](https://developers.google.com/apps-script/guides/sheets).

In the latter case,
- Name the HTML file together with its directory `src/`, i.e, `dialog.html` should be save as `src/dialog.html` on the script editor. You wouldn't need to change the default file name `code.gs`; just copy the contents of `xml2sheets.js` into it.
- Copy the contents of worksheets `Template - Channel` and `Template - Items` to your spreadsheet.  

Be sure to reload the spreadsheet after saving the script.

## How to Use
1. From the menu, select `WordPress XML to Sheets` > `Parse XML`.
1. You will be prompted to upload the XML file that you [exported from your WordPress website](https://wordpress.org/support/article/tools-export-screen/). Select your file and `Upload`.
1. After about a minute of waiting, you will see a couple of worksheets created on your spreadsheet. One to give you a summary of the website and the other for the whole list of contents that you exported.

## Terms and Conditions
You must agree to the [Terms and Conditions](https://ttsukagoshi.github.io/scriptable-assets/terms-and-conditions/) to use this solution.
