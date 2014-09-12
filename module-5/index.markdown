---
layout: post
title: Module 5
---

<div class="objective">
Learning Objective: Importing and cleaning data into a workable format in Excel
</div>

## Importing Sunrise and Sunset Times
In this section we calculate the sunrise and sunset times for every day of the year for Sydney.
Later we'll be able to compare these times with average maximum daily temperatures to see how day length affects temperature.

For now, we're faced with the challenge of sourcing the sunrise and sunset data and importing it into Excel.
One good source of such data is Geoscience Australia's [Astronomical Information](http://www.ga.gov.au/earth-monitoring/astronomical-information.html) service.
This service will compute a variety of astronomical events, including moon phases and planet rise and set information.
They also provide a handy service to [Compute Sunrise, Sunset & Twilight Times](http://www.ga.gov.au/geodesy/astro/sunrise.jsp).
Sunrise, sunset, and civil-, nautical- and astronomical- twilight all have precise [definitions](http://www.ga.gov.au/earth-monitoring/astronomical-information/astronomical-definitions.html).
We'll be using sunrise and sunset times in this course, but sourcing data for these other astronomical events is basically the same.

The data generated is well formed and easy to interpret. 
Here's an excerpt showing sunrise and sunset times for the first week for the first six months of 2014:

![Results format](04.png)

While this table is quite easy for humans to interpet, getting it into a format we can work with in Excel is more challenging.
As is so often the case, getting the data into a format we can work with is most of the challenge.
So a key component of this module is getting this data into a workable format.
At the same time, we want to do this with the minimum of manual effort.
Retyping the whole table is, of course, a possibility. But what an agonising and time consuming task it would be!

We'll of course do things a better way that, while still involving copying-and-pasting, requires no retyping.
Of course, before starting out it's worth having the objective in mind.

In this case we'd like to arrive at an Excel worksheet with 365 rows of data (or 366 for leap years).
Each row would contain the day of the year and the sunrise and sunset time.
In the next module, we'll work on changing the times so they appear and *behave* like times in Excel, rather than plain numbers like ```0448``` and ```1909``` per the original format.

Once the data is in that format, we can do pretty much anything we want with it in terms of analysis.

## Preparation
1. Open **Microsoft Excel** with a new blank workbook.
2. Save the new workbook in your *project directory* as "GASunriseSunsetData.xlsx".
3. Rename the first worksheet "Raw".
3. Add or rename a second worksheet "Processed".
3. Add or rename a third worksheet "Metadata".

# Exercise 1

## Sourcing the Sunrise Sunset Times

### Steps

4. Open a web browser and visit Geoscience Australia's [Compute Sunrise, Sunset & Twilight Times](http://www.ga.gov.au/geodesy/astro/sunrise.jsp) page.
5. Under the section of the page entitled **"Use the National Gazetteer of Australia to determine the Longitude and Latitude"**, enter:
  * A **Place Name** of "Sydney".
  * Select the **Place Type** "Towns & Localities".
  * Select the **State/Authority** "New South Wales".

    The form should appear as follows:
    ![Place name look up](01.png)

6. Click the **Submit** button.
7. The form will refresh and you'll be able to select from a list of matching place names:
    ![Select the place name](02.png)
8. **Click on** the **third location** in the list, labelled **"Sydney"** and at **Lat** -33&deg; 52', **Lon** 151&deg; 12'. The chosen location will automatically be transferred to the second form further down the page:
    ![Sunrise sunset calculator options](03.png)
9. You will also need to:
  * Select **"Defined Event"** and **"Sunrise/Sunset"** under **Select the Required Event**.
  * The current year (_2014_, _2015_, _2016_, etc.) under **Enter Date**.
10. **Click** the **Compute** button. In a few moments, the page will refresh to show "Sunrise & Sunset Results":
    ![Results format](04.png)

    In order to import the data into Excel we need to copy it to the clipboard.

11. Start by using your mouse to select all the data, i.e.:
  * Start your selection at `SYDNEY`
  * Finish your selection at `version 2.2`

    The selection should appear as follows:

    ![Selecting the results](05.png)

12. Now copy the selection to the clipboard, using the **Copy** option in the **Edit** menu. Alternatively you can use the keyboard shortcut `Control-C`.

<div class="note">
  You have successfully sourced your dataset
</div>

# Exercise 2

## Importing into Excel

### Steps

1. Switch to Excel and ensure the "Raw" worksheet is selected in your new workbook.
2. Position the cursor in the *top-left* cell ```A1```.
3. Select **Paste** from the **Edit** menu, or use the keyboard shortcut `Control-V`.

    The Excel worksheet will now appear as follows:

    ![Data pasted into Excel](06.png)

    The data has pasted successfully, but if you look closely it only occupies the first column.
    We'll need to split this data into columns. But first, we need to do a preliminary tidy up.

## Tidying up the data
The first thing to note is that the data contains quite a few blank lines. Fortunately there's a quick way to remove these.

### Next Steps

1. **Choose "Go To..."** from the **"Edit"** menu (on a Mac) or (on a PC) from the **Home** menu **select "Find & Go"** and then **select "Go > To"**.
2. Select the "Special" button from the resulting dialog box.
3. Choose the "Blanks" option and click the "OK" button.

    This will select all blank rows in the worksheet:

    ![All blank rows selected](08.png)

4. Now choose "Delete..." from the "Edit" menu (on a Mac) or (on a PC
) from the **Home** menu **select Delete**.
5. Choose "Entire row" from the resulting dialog box.
6. Click the "OK" button.

    The blank rows will be removed:

    ![All blank rows selected](09.png)

<div class="note">
  You have now imported and tidied your data in Excel.
</div>

# Exercise 3

## Managing the metadata

The first few lines of the data and the last line actually constitute **metadata**:

![Some metadata](10.png)

They tell us about the data, but they're not the data itself.
The give away here is that the metadata is not formatted into columns like the rest of the data.
The fact that it is in a different format can upset the import process, so we need to remove it.

But we don't just want to trash it. Metadata is often a useful reference, particularly 6 months or 6 years down the track when the details are forgotten.
So we'll copy it into the "Metadata" Excel worksheet we created earlier.

### Steps

1. Select the metadata from the top of the "Raw" worksheet:

    ![Metadata preamble](11.png)

2. Cut the cells using the "Cut" option from the "Edit" menu (on a Mac) or (on a PC) from the **"Home"** menu **select** the **"scissors"** icon. Alternatively you can use the `Control-X` keyboard shortcut.
3. Switch to the "Metadata" worksheet.
4. Position the cursor in the top-left cell ```A1```.
5. Paste in the metadata using the "Paste" option from the "Edit" menu (on a Mac) or (on a PC) from the **"Home"** menu **select "Paste"**. Alternatively, you can use the `Control-V` keyboard shortcut.

6. Switch back to the "Raw" worksheet and select the last line of the dataset:

    ![Metadata postamble](12.png)

7. Repeat the process above to paste this into the "Metadata" worksheet.

   Once done your "Metadata" worksheet should look like this:

   ![Metadata postamble](13.png)

Now, if ever in future we want to know what the data is, we can switch to the "Metadata" worksheet and remind ourselves.
Feel free to make additions to the metadata - the more information recorded, the better chance of reusing the data in future.
It would make sense, for instance, to record the link to the service used to calculate the sunrise and sunset times:

  * [http://www.ga.gov.au/geodesy/astro/sunrise.jsp](http://www.ga.gov.au/geodesy/astro/sunrise.jsp)

## More blank rows...

Returning to the "Raw" worksheet, we now have a few blank rows at the top:

   ![Blank rows at top](14.png)

Use the technique described above (or any other technique you know) to remove them.

Your worksheet will now be quite neatly structured:

![Neatly structured data](15.png)

It is now ready to be split into columns.

<div class="note">
  You have now removed extraneous information and rows from your data
</div>

# Exercise 4

## Splitting the data into columns

To split the data into columns, we can use the "Text to Columns" option on the "Data" ribbon:

![Data ribbon - Text to Columns](07.png)

### Steps

To use the "Text to Columns" feature:

1. Select the entire dataset. Since the data only occupies column A, you can accomplish this by double-clicking the top of the ```A``` column:

    ![Select entire dataset](16.png)

1. Select the "Data" ribbon and choose "Text to Columns". The following dialog box will appear:

    ![Text to Columns - Dialog 1](17.png)

1. Choose the "Fixed Width" option and click the "Next &gt;" button. The following dialog box will appear:

    ![Text to Columns - Dialog 2](18.png)

    Excel should correctly infer where each column should begin. Note: it will probably split the month names in half, but that is inevitable since there are fewer month names than there are columns of data.
    The main thing to observe is that there is a split between each time value.

1. Click the "Next &gt;" button. The following dialog box will be displayed:

    ![Text to Columns - Dialog 3](19.png)

1. For each column in the "Data preview", click it and then choose "Text" as the "Column data format":

    ![Text to Columns - Dialog 4](20.png)

    In the picture above, the first four columns have been given the "Text" data format.

1. To complete the import action, click the "Finish" button.

The data will now be split into columns:

![Data split into columns](21.png)

<div class="note">
  You have now split your data across multiple columns
</div>

# Exercise 5

## Unpacking the data

The data is now nicely split it to columns, but the format is still a bit arcane to work with. 
To get the data into a workable format, we need to extract each month's worth of data in sequence and paste it into a different worksheet.
That's what we'll use the "Processed" worksheet for.

We can accomplish this with 12 copy-and-paste actions from the "Raw" worksheet to the "Processed" worksheet. 
Let's run through the process for January:

### Steps

1. Switch to the "Raw" worksheet and select the sunrise and sunset times for January. Since January has 31 days, there should be 31 rows by 2 columns selected:

    ![Selecting January](22.png)

    (Here we've shaded each alternate month's data to make it easier to distinguish.

1. Select "Copy" from the "Edit" menu, or use the `Control-C` keyboard shortcut.

1. Switch to the "Processed" worksheet and paste the data starting at cell "B2". To paste, select "Paste" from the "Edit" menu or use `Control-V`.

    The "Processed" worksheet should look like this:

    ![Selecting January](23.png)

1. Return to the "Raw" worksheet and select and copy the data for February.

1. Paste the data for February below the data for January in the "Processed" worksheet.

1. Repeat the same process for the remaining 10 months.

Once data for all 12 months has been copied across, the last record should be on row 366. This makes sense since there are 365 days in a year and we've left one row at the top of the worksheet blank:

![366 rows](24.png)

<div class="note">
  Your data is now in a workable format.
</div>

# Exercise 6

## Adding a column for the date

```Column A``` in the "Processed" worksheet was left blank. Why? In order to make room for a date column that re-associates each sunrise and sunset pair with its corresponding date.
We know the records are in sequence, with the first one referring to the 1 January of the year (say *2014*). So re-introducing the dates should be quite straightforward.

### Steps

1. The first step is to select the whole of ```column A``` and set its data type to "Date".

1. Next let's give the columns some sensible names, by setting:
  * `A1` to be "Date"
  * `A2` to be "Sunrise"
  * `A3` to be "Sunset"

    You may wish to bold this header row or shade it to make it stand out.

2. Next enter `1/01/2014` in cell "A2" (substitute 2014 with the year to which your data refers).

    That takes care of the first record. Now only 364 to go! Luckily we can use a simple formula to calculate the subsequent dates.

3. In cell "A3" type `=A2 + 1`. This will add one day to the date in "A2", resulting in "2/01/2014".

4. What we want to do is repeat this formula for the remaining cells in ```column A```.  The quickest way to do this to mouse over the bottom-right of A3 until a black cross appears. When it does, double-click and the formula will be copied down ```column A``` until the end of the data:


![Quickly replicating the formula](25.png)

Quickly check your work by verifying that the last line of data corresponds to the last day of the year (31/12/*yyyy*):

![Check your work](26.png)

<div class="note">
Your sunrise and sunset data now has dates associated
</div>

Congratulations! You've now successfully imported a dataset that presented some real challenges. 
The multi-column, space delimited format seemed like something Excel would struggle to import, but it is possible.

You should now be in a good place to import all sorts of data in all sorts of formats. 
It's always worth pausing, however, to consider whether to retype the data.
It's a trade off between time spent working out how to use the various import mechanisms Excel offers versus brute force typing.

Both carry the risk of introducing errors. And for this reason it is important to sanity check your data, as we will do in the next module.

<a class="next-link" href="{{ site.baseurl }}/module-6/">Go to the next module</a>
