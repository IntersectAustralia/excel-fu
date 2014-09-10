---
layout: post
title: Module 7
---
## Putting it all together
In this module we combine our average daily maximum temperature data with our sunrise and sunset data to explore the phenomenon of seaonal lag in coastal Sydney.

## Preparation
1. Open the "BOMWeatherData.xlsx" file you worked with in module 1 and 2 in **Microsoft Excel**.
1. Open the "GASunriseSunsetData.xlsx" file you worked with in module 3 and 4 in **Microsoft Excel**.
1. Open a new workbook in **Microsoft Excel** and save it as "SeasonalLag.xlsx".
1. In the new "SeasonalLag.xlsx" file, rename the first worksheet to "Seasonal Lag"
1. Remove any other (blank) worksheets in "SeasonalLag.xlsx".


## Bringing our data in
First we need to get our data into the same place:

1. Ensure that "BOMWeatherData.xlsx" is the active workbook.

1. Ensure that "Final" is the active worksheet.

2. Select "Move or Copy Sheet..." from the "Edit" menu.

3. In the dialog box that appears, select "SeasonalLag.xlsx" under "To book". Select "Seasonal Lag" under "Before sheet" and ensure that "Create a copy" is checked:

    ![Copy sheets](01.png)

4. Click the ```OK``` button. A copy of the "Final" worksheet should now be present in "SeasonalLag.xlsx".

5. You can now close the "BOMWeatherData.xlsx" workbook.

Now:

1. Make "GASunriseSunsetData.xlsx" the active workbook.

1. Ensure that "Processed" is the active worksheet.

2. Select "Move or Copy Sheet..." from the "Edit" menu.

3. In the dialog box that appears, select "SeasonalLag.xlsx" under "To book". Select "Seasonal Lag" under "Before sheet" and ensure that "Create a copy" is checked:

    ![Copy sheets](02.png)

4. Click the ```OK``` button. A copy of the "Processed" worksheet should now be present in "SeasonalLag.xlsx".

5. You can now close the "GASunriseSunsetData.xlsx" workbook.

<div class="warning">
Note that the worksheets are copied rather than linked. So if you make a change in, say, "BOMWeatherData.xlsx", it will not be reflected in "SeasonalLag.xlsx".
</div>

## Presenting the data on a common plot

A neat thing about copying the worksheets is that our named ranges have also copied across.
If we select "Name > Define..." from the "Insert" menu, we can see all the named ranges defined for our workbook.
For "SeasonalLag.xlsx" it will look as follows:

![Defined named ranges](03.png)

(Click ```Close``` to exit this dialog box).

The named ranges, once again, greatly simplify plotting the data:

1. Ensure the "Seasonal Lag" worksheet is active.

2. Select "Smooth Marked Scatter" from the "Scatter" option on the "Charts" ribbon:

    ![Smooth line scatter](04.png)

3. Choose the "Select" (Data) option from the "Charts" ribbon: 

    ![Select Data](05.png)

4. In the resulting dialog box, repeatedly press the Remove button to remove all series. This is so we start with a blank slate:

    ![Clean slate](06.png)

5. Next click the ```Add``` button.
6. In the "Name" text box enter ```=Processed!J1```.
7. In the "X values" text box enter ```=Processed!EventDate```.
8. In the "Y values" text box enter ```=Processed!DayLength```.
5. Click the ```Add``` button again.
1. In the "Name" text box enter ```=Final!A1```.
1. In the "X values" text box enter ```=Final!WeatherDate```.
1. In the "Y values" text box enter ```=Final!AverageMaxTemp```.
1. Click the ```OK``` button.

A plot will be displayed, but it will be all wrong. Why?

![Bad plot](07.png)

To fix this we need to place the temperature data on a **secondary axis**:

1. Select the data by clicking on the line:

    ![Select data](08.png)

2. Right-click and select "Format Data Series..." from the context menu:

    ![Format data series](09.png)

3. Select "Secondary axis" and click the ```OK``` button:

    ![Secondary axis option](10.png)

The plot makes much more sense now:

![Improved plot](11.png)

But it could still be improved by:

1. Increasing the size of the plot.
1. Re-positioning the legend so it takes up less space.
1. Changing the scales, so less of the plot area is wasted white space.
1. Removing the markers (or making them smaller.)
1. Making the lines thinner to show more detail.
1. Adding a title.
1. Labelling the axes.

Have a go at doing each of these things.

You should arrive at a plot similar to this one:

![Final plot](12.png)

What does this tell us about the relationship between temperature and day length in Sydney?



<a class="next-link" href="{{ site.baseurl }}/module-8/">Where to next?</a>
