# SEC XBRL Data for Excel

A tool for downloading, storing, and querying SEC quarterly and annual XBRL data directly from Excel.

## Why I Built It

The project had two goals:

1. Make SEC XBRL data easier for finance and accounting users to work with through familiar Excel functions and VBA.
2. Explore what scalable, database-backed interactivity from Excel could look like.

Rather than calling a web API every time a user requested data, the tool downloaded SEC datasets, cached them in a local SQL Server database, and exposed low-latency queries through Excel VBA and user-defined functions.

Today, modern analytics platforms and Excel integrations solve much of this much more cleanly. This is not a production scale architecture by any means, but does showcase localized excel based solution capabilities. At the time, however, this was an early exploration of combining Excel’s accessibility with the performance and scale of a relational database.


## How It Works

The intended workflow was:

* Download SEC XBRL datasets
* Create and load SQL Server tables
* Store financial facts locally
* Query those facts from Excel using VBA
* Return results directly into Excel ranges or arrays

Because the data was cached locally, users could repeatedly query large SEC datasets without waiting on external API calls.

## Project Status

* **SEC data download class:** Complete
* **SQL data loading class:** Approximately 50%
* **SQL query class:** Approximately 10%
* **XBRL taxonomy integration:** Not started

## Requirements

* Microsoft SQL Server 2017
* Microsoft OLE DB Driver for SQL Server
* Excel with VBA support

The VBA classes use late binding, so users do not need to manually add library references within the VBA editor.

## Setup

1. Import the class modules into an Excel VBA project.
2. Install SQL Server and the appropriate Microsoft OLE DB driver.
3. Place the `SECVba` project files in the configured local directory.
4. Use the provided class methods to:

   * download SEC data
   * create SQL Server tables
   * load downloaded datasets
   * query financial facts from Excel

Load validation is strongly recommended. For example, the process should record which SEC quarter has already been loaded to prevent duplicate ingestion.

## Example

The example below returns current-asset values for selected SEC filers as of March 31, 2018, directly from SQL Server into an Excel array.

![Example Excel Query Result](https://github.com/dylan1218/SECDataParseVBA/blob/master/ExampleArrayResult.PNG)
