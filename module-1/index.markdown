---
layout: post
title: Module 1
---

<div class="objective">
Learning Objective: Importing data into Excel in a workable format
</div>

## Importing Weather Data 

Being able to locate, download, and work with relevant datasets is a critical research skill. In this module, you will learn how to:

1. locate useful external data sets and source them off the web.
1. download external data and import it into Excel prior to working with it
1. applying formatting to the data
1. enrich the data with meaningful metadata.

<div class="inset">
<div class="title">Where can you find open datasets on the web?</div>

<p>There are plenty of useful datasets available online. Websites such as the <a href="http://australia.gov.au/topics/it-and-communications/data">Australian Government</a> and <a href="http://www.bom.gov.au/climate/data-services/">Bureau of Meteorology</a> provide such datasets. A lot of data sources are becoming freely available on the net. Below are a number of good examples to look at:</p>
<br/>
<ul>
  <li><a href="http://data.gov.au">Australian Government</a> provides an easy way to find, access and reuse public datasets from the Australian Government. </li>
  <li><a href="http://data.nsw.gov.au">NSW Open Data Portal</a> is part of the Open Data initiative of the NSW Government ICT Strategy. NSW has been publishing government data online since 2009. In 2013 a new, upgraded version of data.nsw makes it easier for people to access datasets uploaded by government agencies with advanced search functions. </li>
  <li><a href="http://researchdata.ands.org.au">Research Data Australia</a> is an Internet-based discovery service designed to provide rich connections between data, projects, researchers and institutions, and promote visibility of Australian research data collections in search engines. It is partnering with research institutions and data producing agencies to bring about four transformations to data - unmanaged to managed, disconnected to connected, invisible to findable and single use to reusable - that will enable Australia’s research data to become a national strategic resource to support better, more efficient and defensible research, and improved policy input.</li>
  <li><a href="http://ckan.org/">CKAN (Comprehensive Knowledge Archive Network)</a>: The CKAN site is a registry of open knowledge packages and projects where you can find open knowledge resources or register one of your own. Some examples of available datasets include: Shakespeare’s works, a global population density database, the voting records of MPs, or 30 years of US patents.</li>
  <li><a href="http://www.infochimps.com/">Infochimps.org</a> is attempting to assemble and interconnect the world’s best repository for raw data. “Just as Wikipedia will help you find out something about everything, infochimps.org will help you find out everything about something.”  </li>
  <li><a href="https://www.freebase.com/">Freebase</a> is an open, shared database of the world’s knowledge and it is community built and maintained. Freebase pulls from open data sources like Wikipedia, and MusicBrainz to create structured information on many topics. This site is also easy to navigate and well-designed, which makes it that much better to use.</li>
</ul>
</div>

## Preparation
Create a folder on your system to contain the files you create in this workshop. (We'll refer to this as the **project directory**.)


# Exercise 1



## About this Exercise

In this exercise we will focus on a dataset from the Bureau of Meteorology (BOM). This dataset provides daily weather observations for Sydney, New South Wales. As the first step, you need to get data for the past 3 months from the BOM. Since the BOM data is arranged by month, you need to import 3 separate files from the past 3 months. The data can be accessed at  [Sydney, New South Wales Obeservations](http://www.bom.gov.au/climate/dwo/IDCJDW2124.latest.shtml) page. 

As you can see in figure below, the most recent monthly observation is provided as a table at the beginning of the page. (Note: Because the course material was developed in May 2014, the latest report is shown as "May 2014 Daily Weather Observations").

![Table of weather data](01.png)

If you scroll to the **Other formats** section of the page, you'll see it's possible to download the data as **PDF** and as **plain text**. We'll be using the plain text version, since this will give us raw data we can efficiently work with.

![Other formats](02.png)

If you scroll down the page to the "Other times and other places" section, the most recent 14 months of observations are available for download:

![Other times and places](03.png)



To get the data for the last three months:

## Acquire the Data

### Steps

1. Browse to the BOM page for the [current month](http://www.bom.gov.au/climate/dwo/IDCJDW2124.latest.shtml).

2. Scroll down to the **Other formats** section and right click on the **plain text version** link. Choose the option to "Open the Link in New Tab".  A text file containing data separated by commas will be displayed: 

    ![CSV data](04.png)

3. Save this file to your **project directory**.

4. Once the data is saved, you can close off the tab.

5. Return to the [BOM Daily Weather Observations](http://www.bom.gov.au/climate/dwo/IDCJDW2124.latest.shtml) page, scroll to "Other times and other places" and repeat the process above to download the data for the two prior months. For instance, if the current month is May 2014, download and save the data for March and April 2014 as well.

<div class="note">
You should now have three files in your <b>project directory</b> named similar to:
</div>

![CSV data](05.png)

<div class="inset">
<p>Note how the file names follow a pattern. They all begin with a common prefix <code>IDCJDW2124</code>, followed by a period (<code>.</code>), followed by the year and month (<code>201405</code>), followed by an extension (<code>.csv</code>) </p>

<p>The common prefix is used by the BOM to uniquely identify the data product, i.e. observations for Sydney, NSW. The year and month give us the period to which the data in the file relates. The extension tells us that the file is in <a href="http://en.wikipedia.org/wiki/Comma-separated_values">comma-separated values</a> format (an oft used, but somewhat ambiguously defined standard for exchanging data).</p>
<br/>
<p><a href="http://www.recordsmanagement.ed.ac.uk/InfoStaff/RMstaff/RMprojects/PP/FileNameRules/Rules.htm">File naming conventions</a> as a great way to encode valuable data about a file. Adopting the habit of putting dates into a filename can be very helpful. If dates are encoded in <code>YYYYMMDD</code> format (e.g. “20140529”) and used as the first element in a file name, files sorted by name will be listed in chronological order.</p>
<br/>
<p>Of course the best kinds of naming conventions are documented somewhere. A quick “read me” file listing what the conventions are is invaluable for future reference. <em>Does the date in the file name refer to the date the data was collected or the date it was analysed?</em> Hmmm... better check the “read me”!</p>
</div>

# Exercise 2

## Importing the data into Excel

### Steps
Now we have our data, we can import it into Excel for analysis:

1. Open **Microsoft Excel** with a new blank workbook.
2. Save the workbook in your *project directory* as "BOMDailyData.xlsx".
3. Rename the first worksheet "Raw".
4. Delete any other worksheets (e.g "Sheet1", "Sheet2") if they exist.

5. Select the "Text" option from the "External Data Sources" section of the "Data" ribbon:

    ![Data Ribbon - From Text](06.png)

6. In the file selection dialog box that appears, browse to the project directory and select the earliest file chronologically out of the three CSV data files.
7. Click the ```Get Data``` button.
8. On the resulting screen, ensure that the "Delimited" option is selected and click the ```Next >``` button:

    ![Delimited option](07.png)

9. On the next screen, ensure only the "Comma" option is selected under "Delimiters". Click the ```Finish``` button:

    ![Comma delimiter](08.png)

10. The next dialog box will ask you where the data should be placed. Choose the first cell of the current sheet by selecting the "Existing sheet" option and entering ```=$A$1``` into the text box. Click the ```OK``` button:

    ![Select destination](09.png)

The data will be imported into the "Raw" worksheet:

![Imported data](10.png)

Now repeat the import process for the remaining two files, leaving the newest of the files until last.

In each case, before repeating the import process, select the cell in ```Column A``` of the first completely blank line. That way, your newly imported data will be inserted immediately after the previously imported data. In this case the correct cell is ```A42```:

![First blank cell](11.png)

<div class="note">
You should now have three months of weather data imported into Excel.
</div>

### A Quick Check
Let's quickly check over our imported data.

* First we note that each imported CSV file has about 9 or so lines of metadata at the top:

![Metadata](12.png)

* This is followed by a row of headings for each column of data.
We also know that each month is about 30 days long.  So we'd expect about ***3 x (9 + 1 + 30) = ~120*** rows of data.
<br/><br/>
* Of course the precise figure may be a little bit more (because some months have 31 days) or somewhat less (say, if we're only a little way into the current month). 
<br/><br/>
* But if our row count is not wildly off our estimate, then our confidence in the data is boosted.

### Take a Copy
<div class="inset">
<p>Our data is pretty much untouched from the form delivered to us by the BOM.
To make it truly useful, we need to remove the metadata and duplicated column headings,
so we have a nice continuous presentation of the data.</p>
<br/>
<p>But it would be nice to preserve the metadata and the data in the original form, so
that we can refer back to it.
This will help us determine, in future, if an unexpected value is the result of our 
processing or if it occurs in the original data provided to us.</p>
<br/>
<p>So we’ll copy the entire “Raw” worksheet to a worksheet we’ll name “Processed”. We’ll do all our processing, formatting, and analysis in the “Processed” sheet, leaving the “Raw” sheet for posterity.</p>
</div>

### Steps

To copy the worksheet:

1. Select "Move or Copy Sheet" from the  "Edit" menu.

2. In the dialog box that appears, ensure "BOMDailyData.xlsx" is selected under "To book". Ensure "(move to end)" is selected under "Before sheet". And finally ensure that "Create a copy" is checked. Once these settings are made, click the ```OK``` button:

    ![Move or Copy sheet](13.png)

3. Rename the copied sheet "Processed".

At this stage, we could [protect the "Raw" worksheet](http://office.microsoft.com/en-au/excel-help/password-protect-worksheet-or-workbook-elements-HP010078580.aspx), thereby preventing any accidental changes. This would be a wise move, particularly for data that's hard to re-import.

<div class="note">
  You should now have a copy of your worksheet.
</div>

# Exercise 3

## Cleaning up the data

### Steps

Now that we have a copy of the worksheet, we can clean it up.

1. Ensure the "Processed" worksheet is active.

2. Select the whole of ```Column A``` and delete it. This will remove all our metadata.

3. Now remove the blank lines at the top, so the column headers appear in the first row:

    ![Remove column headers](14.png)

4. Now **scroll down** the page and **remove the blanks and redundant header**:

    ![Remove blanks and redundant header](15.png)

5. And **repeat** once more to **remove the last set of blanks and the last header**.

The result should be a uniform worksheet, with a single set of headers in the first
row and up to 90 or so rows of weather records for the past 3 months:

![Final data file](16.png)

<div class="note">
You should now have your data in a workable format. Congratulations! Getting the data into a format you can use is often 90% of the challenge.
</div>

