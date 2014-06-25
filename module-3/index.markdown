---
layout: post
title: Module 3
---
## Importing Historical Weather Data 

In this module we locate historical weather data for Sydney and import it into Excel. We'll quickly check that the import has worked, but we'll leave detailed exploration of that data for the next module.

## Preparation
1. Open **Microsoft Excel** with a new blank workbook.
2. Save the new workbook in your *project directory* as "BOMWeatherData.xlsx".
3. Rename the first worksheet "Raw".
4. Delete any other worksheets (e.g "Sheet1", "Sheet2") if they exist.

## Acquire the data
The Bureau of Meteorology provides [Past Weather &amp; Climate Data](http://www.bom.gov.au/climate/data-services/).
In this section we're going to import historical temperature data for Sydney.
The [Observatory Hill weather station](http://www.bom.gov.au/clim_data/cdio/metadata/pdf/siteinfo/IDCJMD0040.066062.SiteInfo.pdf) has been collecting data since January 1859.
We're going to get the maximum temperature for every day this weather station has operated.

1. Start by browsing to [Climate Data Online](http://www.bom.gov.au/climate/data/).
2. Select "Temperature" from the ```Data about``` drop down list.
3. Select "Daily" and "Maximum temperature" from ```Type of data```.
4. Enter "Sydney" in the ```Select a weather station in the area of interest``` text box.
5. Click the ```Find``` button:

    ![Climate Data 1](01.png)

6. Select the first entry, "Sydney", from ```Matching towns```. And then select "066062 Sydney (Observatory Hill)..." from ```Nearest Bureau Stations```:

    ![Climate Data 2](02.png)

7. Select the current year (e.g. "2014") in the ```Select Year``` drop down box:

    ![Climate Data 3](03.png)

8. Finally, click the ```Get Data``` button.

A new tab or window will open in your browser. When it does, click the ```All years of data``` link at the top right:

![Climate Data Download](04.png)

The zipped data will download to your computer. (It will be called something like "IDCJAC0010_066062_1800.zip").

Copy this file to your **project directory** and unzip it.

## Importing the Data

Now that we've downloaded our data, we can import it into Excel:

1. Switch back to Excel and ensure the "Raw" worksheet is active.

2. Select the "Data" ribbon.

3. Select "Text" from the "External Data Sources" section of the ribbon.

4. Browse and select the CSV file of weather data from your projects directory. Click the ```Get Data``` button.

5. Select "Delimited" as the "Original data type" and click the ```Next``` button:

    ![Delimited](05.png)

6. Ensure "Comma" is (exclusively) selected as the delimiter. Quickly check the "Data preview" pane to ensure the data is being correctly broken up into columns:

    ![Delimited](06.png)

7. Click the ```Finish``` button.

8. On the resulting dialog, ensure the data will be placed in cell ```$A$1``` and click the ```OK``` button:

   ![Ensure the data will be placed in cell $A$1](07.png)

## Quick Checks

The imported data should look as follows:

![Imported BOM data](08.png)

There is a lot of redundancy in this data. The "Product code" and "Bureau of Meteorology station number" are the same for every line.
This is what we'd expect since we asked for one type of data (the "product") for one station ("66062 - Observatory Hill").
Next we have separate columns for "Year", "Month" and "Day". Taken together, these represent the day the temperature was recorded.
When combined they are unique for each line of the file.

Next comes the actual temperature reading ("Maximum temperature (Degree C)").

Lastly there are some additional columns that we won't use for this exercise ("Days of accumulation of maximum temperature" and "Quality").

In terms of columns, the data checks out. What about in terms of row count?

Go to the last row of data. Is the number of rows of data roughly plausible given we'd expect 1 row of data for each day since 1/1/1859?

(As a quick and dirty check, you might like to enter the following formula into Excel to estimate the expected row count ```=(YEAR(TODAY())-1859)*365.25```. It's a bit rough, but it will tell us if we're in the ballpark. How does it work?)


