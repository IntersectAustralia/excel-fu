---
layout: post
title: Module 2
---
## Exploring the Daily Weather Data 

This module aims to help you build on your analytical skills by learning how to explore data in Excel. We will dig into the imported BOM data to examine if the data looks plausible and meaningful. By the end of this module you will be able to apply a number of features such as basic functions, sorting, filtering, conditional formatting, charting, and pivot tables and charts.

## Preparation
1. Open **Microsoft Excel** 
2. Open the "BOMDailyData.xlsx" workbook you created in the previous module.

## Does everything look Okay?

As the first step, we should explore if the weather data looks plausible. Calculating minimums, averages and maximums for maximum temperature, minimum temperature and rainfall is a good place to start.

1. Scroll down and position your cursor in ```Column B``` on the first blank line after the last line of data.

2. Type in ```=MIN(```, then select all the data above in ```Column B```, then type ```)```, then hit the enter key.

The minimum temperature will be calculated for you.

Next, repeat the process to calculate the average in the cell immediately below. Hint, use ```=AVERAGE(...)```.

And repeat the process again to calculate the maximum in the cell immediately below that one. Hint, use ```=MAX(...)```.

Now follow the steps above to calculate the mimimum, average, and maximum for columns ```C``` and ```D```.

When done your sheet will look something like this:

![Min, average, max](03.png)

<div class="tip">
Tip: Sometimes it's nice to see the formulas on a worksheet rather than the calculated values.
Simply use the keyboard shortcut <b>Control-`</b> ("`" is above the tab key) to accomplish this.
</div>

<div class="warning">
Warning: You may notice little green triangles in the top-left of some cells:
These are where Excel has detected a possible inconistency in your worksheet data or formulas. Always check these, if only to ignore them:
</div>

![Validation errors](10.png)

Applying a bit of formatting can really to assist with the interpretation of the figures:

![Min, average, max formatting](04.png)

As can [hiding rows](http://office.microsoft.com/en-au/excel-help/show-or-hide-columns-and-rows-HP010342574.aspx) so we can align the calculated values with the column headings:

![Hiding cells](05.png)

<div class="warning">
Warning: Hiding rows can cause confusion. Be sure to check the row numbers &mdash; a break in the sequence indicates there are hidden cells.
</div>

To get around the issues with hidden rows you may prefer to [freeze panes](http://office.microsoft.com/en-au/excel-help/freeze-or-lock-rows-and-columns-HP001217048.aspx). You can freeze the header row, for instance, which will cause it to remain static at the top of your screen, while all the data rows scroll.

(If you have hidden cells, unhide them before moving to the next section.)

The upshot: the minimums, averages and maximums all look plausible for Sydney for that time of year, so our data is probably sound.

## Coldest nights? Hottest nights?

Conditional formatting is great for highlighting cells based on a given rule. You might for instance highlight cells whose value is above or below a threshold. Or you might want to highlight cells with the ```n``` highest or lowest values.

In this case we want to add two rules. The first to highlight the hottest 10% of nights <span style="color:red">red</span> and the second to highlight the coldest 10% of nights <span style="color:green">green</span>.

To apply conditional formatting to the mimimum temperature:

1. Select the data in ```Column B```.

2. Select "Conditional Formatting" from the "Home" ribbon. Then choose "Top/Bottom Rules..." and "Top 10%...":

    ![Conditional Formatting Top 10%](01.png)

3. On the resulting dialog box, enter "10" in the text box and click the ```OK``` button:

    ![Top 10% Rule](02.png)

The cells which contain the top 10% of minimum nightly temperature (i.e. the hottest nights) should be highlighted red:

![10% Rule Output](06.png)

Now see if you can highlight the coldest 10% of nights green.

For the data in ```Column C```, containing the Maximum Temperature, see if you can display some red "Data Bars":

![Data bars](07.png)

For the rainfall data in ```Column D```, try a "Colour Scale".

Your worksheet should now look something like this:

![Worksheet](08.png)

<div class="tip">
Tip: The "Data Bars" and "Colour Scales" offered by conditional formatting are great for showing differences between values in adjacent cells. What if you wanted to summarise multiple values, graphically in a single cell? <a href="http://www.excel-easy.com/examples/sparklines.html">Sparklines</a> are tiny graphs that occupy a single cell, based on data in a range of cells. They are very cool. Here are some that summarise the minimum and maximum temperatures and the rainfall:
</div>

![Worksheet](09.png)

## Sorting
Our data comes from the BOM sorted by date. Excel allows us to easily sort it by another column. We can even sort data by multiple columns (the classic example being sorting by a _surname_ column, then a _first name_ column). In our case we want to sort by the **Minimum temperature** column.

1. Start by selecting the rows of data, including the header row.

2. Select **Sort** from the **Data** menu.

3. In the resulting dialog box, select "Minimum temperature" as the **Column**, select "Values" for **Sort On**, and select "Smallest to Largest" for the **Order**. Click the ```OK``` button:

    ![Sorting dialog box](11.png)

Our data is now sorted by minimum temperature and our conditional formatting is still active, so the green shaded, bottom 10% of temperatures appear first:

![Sorted data](12.png)

## Filtering
Filtering is a great way to narrow down the data to a subset. If a value is continuous, we can filter between upper and lower limits (e.g. _give me all the rows where this column is between 1.01 and 3.14_). If the data is categorical, (e.g. _ACT, NSW, QLD, Vic,..._) we can select which values to display.

In the case of the daily BOM data, there is a **Direction of maximum wind gust** column. We can filter for all days where the direction of maximum wind gust is westerly (i.e. "W").

To enable filtering:

1. Select **Filter** from the **Data** menu. Little arrows will appear in the header row:

    ![Filtering](13.png)

2. Click on the arrow for the **Direction of maximum wind gust** column.

3. On the pop-up box that appears, select to filter by the "W" value:

    ![Filtering](14.png)

The worksheet will be filtered to show only rows where the wind direction is Westerly:

![Worksheet filtered](15.png)

<div class="tip">
Tip: To clear the filter, select "Filter" again from the "Data" menu.
</div>

Make sure you clear the filter before moving to the next section.

## PivotTables
Filtering is great for viewing a subset of the data. But what if you want to know how many days the maximum wind gust was a Westerly as compared to other wind directions? For this we need to aggregate our data, and a PivotTable is well suited to this task.

To add a PivotTable:

1. Select your data, including the header row.

2. Select **PivotTable...** from the **Data** menu.

3. In the resulting dialog box, ensure "New worksheet" is selected and click the ```OK``` button:

    ![PivotTable dialog](16.png)

4. A new worksheet will be inserted and the **PivotTable Builder** will appear. Configure it by dragging the "Direction of maximum wind gust" field as shown:

    ![PivotTable builder](17.png)

This should generate a PivotTable as follows:

![PivotTable view](18.png)

As you can see, Westerly winds dominated during the period of data coverage for Sydney.

Try playing with the PivotTable to see what other insights into the data you can get.

Also, consider renaming the worksheet. By default it will be called "Sheet1" or similar. "Weather Summary" would be much more meaningful six months down the track!

## Compass Text to Angles: VLOOKUP

What if we want to map the [textual compass points](http://en.wikipedia.org/wiki/Points_of_the_compass) (e.g _N_, _E_, _SSW_) to numeric values representing degrees in a circle (i.e. 0 to 360&deg;)? Here, the ```VLOOKUP``` function is our friend.

Let's get under way:

1. Start by creating a new worksheet and calling it "Compass":

2. Create a table in the "Compass" worksheet as follows:

    ![Compass worksheet](19.png)

    Here is one way to build the table:

    ![Formulas](20.png)

    Can you explain how this works?

3. Now select the data in the table and name this range "CompassPoints":

    ![CompassPoints range](21.png)


Now switch back to the "Weather Summary" worksheet, containing the PivotTable.

1. Locate your cursor in cell ```C5```. This should be immediately to the right of the first line of summarised data.

2. Type in the formula ```=VLOOKUP(A5, CompassPoints, 2, FALSE)``` and press the **Enter** key.

3. Now double-click in the bottom right of the cell to replicate the formula in the cells below:

    ![Replicate formula](22.png)

The ```VLOOKUP``` formula will look up the **CompassPoints** range. It will try to locate a value in the first column of that range that matches the value in cell ```A5```. When it finds that value, it returns the corresponding value in the second column of the **CompassPoints** range, i.e. the value in degrees.

Our table will now contain mapped values in degrees for the textual compass points:

 ![Vlookup result](23.png)

(As ever, it's a good idea to give a name to the column, as I've done above.)

<div class="tip">
  Tip: <b>VLOOKUP</b> is great for these kinds of "translation" activities, i.e.
  mapping one value to another. It's also great for "tax bracket" style problems,  where for instance you want to apply a certain tax rate if the income value
  lies within a specific range.
</div>


## Charting

Later on we will do some X-Y scatter plots to show time series data. For the moment we'll try a [polar style plot](http://en.wikipedia.org/wiki/Polar_coordinate_system). These are particularly useful for showing wind direction. Excel calls these "radar" plots.

To get started:

1. Create a new worksheet and name it "Chart".

2. Copy the data from the "Weather Summary" tab to the clipboard.

3. Paste the data into the "Chart" worksheet using **Paste Special...** from the **Edit** menu.

4. In the dialog box that is displayed select the "Values" option and click the ```OK``` button:

    ![Paste special](24.png)

This will copy the data to the new sheet while discarding all formulas. The result should look something like this:

![Chart worksheet](25.png)

Now:

1. Sort this data by the "Compass Points as Angles" in ascending order (smallest to largest).

2. Rename "Row Labels", which isn't a very meaningful column name, to "Direction".

3. Rename "Total", which also isn't very meaningful, to "Days".

4. Lastly select the headers and data in the "Direction" and "Days" columns:

    ![Direction and Days data](26.png)

With our data selected, we can now add the chart:

1. Select the "Radar" chart type from the "Other" section of the "Insert Chart" portion of the "Charts" ribbon:

    ![Radar chart selection](27.png)

This should generate a new chart as follows:

![Wind gust chart](28.png)

Clearly, the West is the dominant direction of the strongest gusts, followed by the South.

Note the format of this chart could be improved. At the very least a meaningful title would help the reader make sense of it. Can you think of a meaningful title and apply it to this chart?

