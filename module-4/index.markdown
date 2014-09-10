---
layout: post
title: Module 4
---
## Working with the Historical Weather Data

In this module we process and explore the historical weather data we imported in module 3, with the ultimate aim of plotting it. While we could plot every single maximum temperature for every year, we will instead plot the average daily maximum for each day of the year over the period since records began.

## Preparation
1. Open **Microsoft Excel**.
2. Open the file "BOMWeatherData.xlsx" in your **project directory** (this file was created in module 3).
3. Create a copy (duplicate) of the "Raw" worksheet, called "Processed".

To do step 3:

1. Select "Move or Copy Sheet..." from the "Edit" menu.

2. Select the "(move to end)" option.

2. Ensure "Create a copy" is checked.

3. Click the "OK" button:

    ![Copy sheet](01.png)

4. Rename to sheet to "Processed".

## Pivot Tables Revisited

To calculate the average of the maximum temperature for each day of the year, we can use a Pivot Table. We're trying to aggregate some data, so a Pivot Table is the natural choice.

To do so:

1. Ensure you have the "Processed" sheet selected.

2. Delete the entire "Product Code" and "Bureau of Meteorology station number" columns.

3. Delete the entire "Days of accumulation of maximum temperature" and "Quality" columns.

    You should be left with 4 columns of data as follows:

   ![4 columns of weather data](02.png)

4. Select the "Data" section of the ribbon and choose the "PivotTable" option. Excel will automatically generate a PivotTable, but it won't be in the format we require.

To get the right output, we need to configure the PivotTable Builder as follows:

![PivotTable Builder](03.png)

This is telling Excel to format the PivotTable to show the Average Maximum Temperature for every day of the year.
So the first row of data will contain the average of all the January 1sts in the dataset.

This is the output that is produced:

![PivotTable Output](04.png)

To configure the PivotTable take the following steps:

1. Uncheck "Year" under ```Field name```. Ensure the other fields, "Day", "Month", and "Maximum temperature (Degree C)", are checked:

    ![Field selections](05.png)

2. Ensure "Day" and "Month" appear under ```Row Labels```. If they don't you can drag and drop them into this area:

    ![Row labels](06.png)

3. The ```Values``` area will default to "Count of Maximum temperature (Degrees C)". This is not quite what we want. (Try inspecting the PivotTable to see if you can work out what this data pertains to).
   Instead we want the **average** of the maximum temperature. To change from a **count** to an **average** click on the little **_i_** button:

    ![I Button](07.png)

4. Select the "Average" option from the ```Summarize by``` drop down list and click the ```OK``` button:

    ![Change field to Average](08.png)

The PivotTable should now appear as in the screen shot above.

You'll note that the PivotTable has been created in a new worksheet. Rename this "Average" and move it to the right of the other worksheets:

![Rename, reorder sheets](09.png)

## Getting the data ready for plotting

The PivotTable is looking good, but some of its current settings will upset a plot. Most importantly, we want to hide all the various totals, since we want to plot just the averages for each day.

To do so:

1. Position the cursor within the PivotTable.

1. Select "Layout &gt; Tabular Layout" from the ```Data > PivotTable``` ribbon:

    ![Tabular View](10.png)

1. Uncheck  "Show for Rows" and "Show for Columns" from the "Totals" option of the ```Data > PivotTable``` ribbon:

    ![Totals](11.png)

1. Select the first "Total" line. This should be in cell ```A36```. Right-click and select "Field settings" from the menu.

1. In the resulting dialog box, select the "None" option under "Subtotals" and click ```OK```:

    ![Remove subtotals](13.png)

This last step will remove the subtotals, leaving just the aggregated data for each day of the year:

![Subtotals removed](14.png)

## Flattening the PivotTable
We've simplified our pivot table by removing all totals. Now we want to flatten it so every row contains the value of the month (instead of blanks).

To accomplish this:

1. Create a new worksheet and name it "Final".

2. Return to the "Average" worksheet, which features your PivotTable.

1. Select the "Select" option from the "Data" section of the ```Data > PivotTable``` ribbon and choose the "Entire Table" option:
    
    ![Selecting Table](15.png)

1. Copy the selected cells. (The quickest way is with the ```Control-C``` keyboard shortcut).

2. Switch to the "Final" worksheet.

4. Position the cursor in cell ```A1```.

5. Go to "Paste Special..." in the "Edit" menu.

6. Select the Paste "Values" option and click the ```OK``` button:

    ![Paste values](16.png)

Pasting values will paste the computed data from our PivotTable, removing any formulas that may be present.

Your "Final" worksheet should now look as follows:

![Final worksheet](17.png)

Note how the "Month" column is mostly blank. In order to completely flatten our data, we need to fix this. The easiest way to do this is as follows:

1. Select ```Column A```.

2. Choose "Go To..." from the "Edit" menu.

3. Click the "Special..." button on the resulting dialog box.

4. Select the "Blanks" option and click the ```OK``` button:

![Goto dialog](18.png)

This will select all the blanks in ```Column A```, starting at cell ```A4```.

At this point the cursor should be situated in cell ```A4```.

Now:

1. Type ```=A3``` into cell ```A4```.

2. Then press the ```Control``` and ```Enter``` keys together.

The worksheet should now look as follows:

![Updated worksheet](19.png)

## Turning month and day into a date
Next we need to turn the month and day into a date that Excel can understand. We can do this using the ```DATE``` function, which takes the _day_, _month_ and _year_ parts of a date as parameters and returns a proper date object.

To do this:

1. Insert a new column between columns ```B``` and ```C```.

2. In cell ```C3``` type the formula ```=DATE(2014, A3, B3)```.

3. Replicate the formula in ```C3``` for each row of data.

4. Add a meaningful heading in ```C2```, such as "Weather Date".

Your worksheet should now look as follows:

![Updated worksheet](20.png)

## Name your Ranges
Now that our data is in a format we can use, we can plot our data.

This is made easier by creating **named ranges** for our key data.
Named ranges let you refer to one or more cells using a meaningful name.
We know, for instance, that our dates occupy the cell range ```C3:C368```.
We can associate the name ```WeatherDate``` with these cells.

Then whenever we need to use the range, we can quickly and meaningfully reference it by its name.
So we can find out the latest date we have coverage for by entering the formula ```=MAX(WeatherDate)```.

This is so much more meaningful, when you revisit the worksheet a year down the track, than trying to work out what ```=MAX(C3:C368)``` is calculating.

To create a named range for the "Date" column:

1. Select the data in column ```C``` (e.g. ```C3:C368```).  The quickest way to do this is place the cursor in cell ```C3``` and press the ```Control-Shift-Down``` keyboard shortcut (or ```Command-Shift-Down``` on the Mac).

2. In the name box, type "WeatherDate" and press the Enter key:

   ![Named "WeatherDate" range](21.png)

   Now repeat steps 1 and 2 to create a named range called ```AverageMaxTemp```, based on the data in column ```D``` (for cells ```D3:D368```).

## Plotting the Data
Because our data is timeseries data, it makes sense to plot it on a scatterplot. The ```x``` axis will be the weather date
and the average maximum temperature will be plotted on the ```y``` axis.

Just as our named ranges made it easy to apply formulas to the data, they will simplify selecting the data to plot.

Let's get under way:

1. Select "Smooth Lined Scatter" from the "Scatter" option on the "Charts" ribbon:

    ![Smooth Lined Scatter](22.png)

2. Choose the "Select" (Data) option from the "Charts" ribbon:
    ![Smooth Lined Scatter](23.png)

3. In the resulting dialog box, repeatedly press the ```Remove``` button to remove all series. This is so we start with a blank slate:

    ![Clean slate](24.png)

4. Next click the ```Add``` button.
5. In the "Name" text box enter ```=Final!A1```.
6. In the "X values" text box enter ```=Final!WeatherDate```.
7. In the "Y values" text box enter ```=Final!AverageMaxTemp```.
8. Click the ```OK``` button to close the dialog box.

After resizing, your plot will look similar to this one:

![Series](25.png)

The final step, to make this graph really useful for publication or for later reference
is to add ```x``` and ```y``` axis labels (including units). In case you've forgotten how,
[here are some instructions](http://office.microsoft.com/en-au/excel-help/add-or-remove-titles-in-a-chart-HP001234162.aspx).
In this case, it would make sense to hide the **legend** since it reduces the size of the plot while adding no additional information.
Finally, it would be worth tweaking the scale. Excel auto-scales the ```y``` axis, but for some reason starts it at zero. 
The minimum average maximum is above 15&deg;C, so we'd see more detail if the scale started at 15.

Similarly, a thinner line can reveal more of the detail:

![Final plot](26.png)

The output is as we'd expect. The average of the maximum temperature peaks in the height of summer (January) and falls to a minimum in the dead of Winter (July).
While the curve is quite smooth and approximately resembles a sine wave, it's intriguing to note the day-to-day variation in temperature, even when averaged over 150 years.

<div class="warning">
Note: a slight problem exists with this plot that could confuse interpretation. Can you spot it?
</div>

<a class="next-link" href="{{ site.baseurl }}/module-5/">Go to the next module</a>
