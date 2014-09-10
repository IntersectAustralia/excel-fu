---
layout: post
title: Module 6
---

<div class="objective">
Learning Objective: Working with and visualising the data
</div>

## Working with the Sunrise and Sunset Data

In module 5 we got to the point where we had an Excel worksheet that looked like this:

![Where we got to](01.png)

It had 3 columns, one consisting of the day of the year, another of the sunrise time and another with the sunset time.
It had 365 rows (or 366 if the chosen year was a leap year).

While the data looks sound, a couple of issues remain:

* the sunrise and sunset times are represented as strings. We can't do much with these until we get them into a time format.
* we haven't done any proper checks to make sure the data has copied across correctly. What if, in all our copying and pasting, we pasted July's data where June's data should have gone?

In this module we'll address these issues.

## Preparation
1. Reopen the "GASunriseSunsetData.xlsx" file you created in module 5 in **Microsoft Excel**.
2. Ensure that you have the second worksheet, labelled "Processed", opened.

## Strings to Times

In computer speak a bunch of characters is called a [string](http://en.wikipedia.org/wiki/String_\(computer_science\)).
In Excel strings have the data type "Text". Strings can contain any characters, including numbers (0-9), letters (A-Z, a-z) and punctuation.
In order to make sense of these strings, we need to understand the format.

In the case of our sunrise and sunset times, the format is ```HHMM```, i.e.
each string is exactly four characters long. 
The first two characters contain the hours component of the time. i.e. ```HH``` could range from _00_ (midnight) to _23_ (11pm).
The second two characters contain the minutes component of the time. i.e. ```MM``` could range from _00_ to _59_, representing the 60 minutes that make up an hour.

Given this knowledge, we as humans can make sense of these times. e.g. we can translate _0448_ to _4:48am_.
Excel, however, only understands _0448_ to be a string. We need to explicitly tell it that it's a time.

So often in life, it makes sense to tackle a bigger problem by breaking it up into a number of smaller problems.
So too with converting this string to a time Excel understands.

We know that each string actually contains two bits of information - the hour and the minute. 
So a logical place to begin is to split the string into these two parts.
Once we've separated our hour information from our minute information, it becomes far easier to recombine them into a time format Excel can work with.

Fortunately Excel provides some tools, in the form of **functions**, that help us split strings into parts.

Excel's functions are analogous to [mathematical functions](http://en.wikipedia.org/wiki/Function_\(mathematics\)). 
You may recall the notation:

```y = f(x)```

The function ```f``` takes some input ```x```, and generates the output ```y```. 
So if ```f = x^2```, then for ```x = 2```, ```y = f(x) = x^2 = 2^2 = 4```.

Excel has a bunch of [functions for working with strings](http://office.microsoft.com/en-au/excel-help/text-functions-reference-HP010079191.aspx).
One is called [LEFT](http://office.microsoft.com/en-au/excel-help/left-leftb-functions-HP010062568.aspx?CTT=5&origin=HP010079191).
Another is called [RIGHT](http://office.microsoft.com/en-au/excel-help/redir/HP010062576.aspx?CTT=5&origin=HP010079191).

```LEFT``` takes two inputs. The first is a string to apply the function to; the second is a number ```n```. 
Whatever ```n``` is, ```LEFT``` will return that many characters of the input string.
So if the string is ```HELLO``` and the number is ```2```, ```LEFT``` will return ```HE```.
As you might already have guessed, ```RIGHT``` functions in the same way, but returns the right-most ```n``` characters.

Let's first extract the hours information for the sunrise:

1. Position the cursor in cell ```D2```. 
2. Type ```=LEFT(B2, 2)``` and press the **Enter** key.

   ![LEFT on sunrise](02.png)

   ```D2``` will now show the hours:

   ![LEFT on sunrise](03.png)

3. Now copy the formula in ```D2``` to all the cells which follow in column ```D```. The quickest way to do this to mouse over the bottom-right of ```D2``` until a black cross appears. When it does, double-click and the formula will be copied down column ```D``` until the end of the data:

   ![LEFT on sunrise](04.png)

4. Apply the same technique used in steps 1 to 3 to extract the minutes information for the sunrise into ```column E```. Hint: In ```E2``` type ```=RIGHT(B2, 2)```. Then replicate the formula down column ```E```:

   ![RIGHT on sunrise](05.png)

5. Now extract the hours and minutes information for the sunset into the ```F``` and ```G``` columns. At this juncture, it is also prudent to give the new columns some sensible headings (_Sunrise Hour_, _Sunrise Minute_, _Sunset Hour_, _Sunset Minute_):

   ![Time values extracted](06.png)

Now that our times are broken up into their simplest parts, we can recombine them into times Excel understands.
For this we can use the [TIME](http://office.microsoft.com/en-au/excel-help/time-function-HP010342955.aspx?CTT=5&origin=HP010342402) function.
It takes three input parameters, ```hour```, ```minute```, ```second``` and returns an Excel time value.

Times in Excel are actually stored as decimal values that represent fractions of a day.
For example, ```0``` represents midnight, i.e. the beginning of the day; ```0.5``` represents noon or half-way through the day;
```0.75``` represents 6pm, etc. This may seem odd, but it makes finding the difference between any two times as simple as subtracting one decimal number from another.

So the ```TIME``` function returns a decimal number corresponding to the time parts specified as input parameters.

Let's apply it to the sunrise time:

1. In ```H2``` type ```=TIME(D2,E2,0)```. Note how we use ```0``` for the ```seconds``` parameter to ```TIME``` because our sunrise time is not calculated to the precise second. Press the **Enter** key:

   ![TIME function](07.png)

2. Replicate the formula in column ```H``` for the remaining rows:
   
   ![TIME function](08.png)

3. Now repeat steps 1 and 2 in order to put the sunset times into column ```I```. While you're at it, add some sensible headers to columns ```H``` and ```I``` (e.g. _Sunrise Time_, _Sunset Time_):

   ![TIME function](09.png)

Note, we said that Excel stores times as a decimal number. Why then is the time shown as "4:48am"?
Because Excel has detected the time value, it has formatted the cell accordingly.
If we change the data type for the cells back to "General", the times will be displayed as decimal values:

   ![Time as a decimal value](10.png)

<div class="warning">
Warning: If you format the times as "General", make sure you <i>undo</i> before moving on to the next steps.
</div>

## Calculate Day Length
So we've used the ```TIME``` function to create dates Excel can understand. 
This makes it trivial for us to compute the length of each day of the year.
We need only subtract the sunrise time from the sunset time.

To do so:

1. Position your cursor in ```J2```.
2. Type the formula ```=I2-H2``` and press the Enter key.
3. Now replicate the same formula in column ```J``` for each row of data.
4. Finally give the column a heading "Day Length". 

You should end up with a worksheet that looks like this:

![Day length added](12.png)

We can instantly see that the day length on New Year's Day is 14h 21m, roughly as we'd expect given the early sunrise and late sunset in summer-time Sydney.

## Name your Ranges
Now that our data is in a format we can use, we can start exploring, sanity checking, and using it.

This is made easier by creating **named ranges** for our key data.
Named ranges let you refer to one or more cells using a meaningful name.
We know, for instance, that our dates occupy the cell range ```A2:A366```.
We can associate the name ```EventDate``` with these cells.

Then whenever we need to use the range, we can quickly and meaningfully reference it by its name.
So we can find out the latest date we have coverage for by entering the formula ```=MAX(EventDate)```.

This is so much more meaningful, a year down the track, than trying to work out what ```=MAX(A2:A366)``` is calculating.

To create a named range for the Date column:

1. Select the data in column ```A``` (e.g. ```A2:A366```). The quickest way to do this is place the cursor in cell ```A2``` and press the ```Control-Shift-Down``` (```Command-Shift-Down``` for Mac users) keyboard shortcut.

2. In the name box, type "EventDate" and press the Enter key:

   ![Named "Date" range](11.png)

Now repeat steps 1 and 2 to create the named ranges ```SunriseTime```, ```SunsetTime``` and ```DayLength```, based on the data in columns ```H```,```I``` and ```J```.

## Sanity Checking the Data
We can now run some checks to see if the data fits with our intuitive understanding of the seasons:

1. Locate a blank cell, say ```L2``` and enter the formula ```=MIN(EventDate)```.
1. Locate a blank cell, say ```L3``` and enter the formula ```=MAX(EventDate)```.

Note, if you see output like ```41640``` and ```42004```, be sure to format these cells as Date.
Just as Excel stores times as a decimal value between 0 and 1, it stores dates as whole numbers.
(And combined dates and times, you guessed it, are stored as a number with a whole and factional part.)

With any luck the minimum and maximum dates will reflect the first and last day of the year to which the data pertains.

Next have a go at:

1. Calculating the earliest sunrise time in ```L4```
1. Calculating the latest sunrise time in ```L5```
1. Calculating the earliest sunset time in ```L6```
1. Calculating the latest sunset time in ```L7```
1. Calculating the duration of the shortest day in ```L8```
1. Calculating the duration of the longest day in ```L9```

Again, you may need to format these cells as Times.

To make your worksheet easier to follow, annotate your calculations in ```K2``` to ```K8```.
You should end up with a worksheet that looks something like this:

![Worksheet with calculations](13.png)

For extra credit, you might also like to calculate the average day length across the year.
Is the figure exactly what you'd expect? If not, can you explain the discrepancy?

## Plotting the Data
These quick checks suggest the data is good. But nothing beats plotting the data to
point out issues with it. Gaps, outliers and data that's out of sequence become immediately apparent when data is plotted.

Because our data is timeseries data, it makes sense to plot it on a scatterplot. The ```x``` axis will be the event date
and the specific event (sunset, sunrise, day length) will be plotted on the ```y``` axis.

Just as our named ranges made it easy to apply formulas to the data, they will simplify selecting the data to plot.

Let's get under way:

1. Select "Smooth Lined Scatter" from the "Scatter" option on the "Charts" ribbon:

    ![Smooth Lined Scatter](14.png)

2. Choose the "Select" (Data) option from the "Charts" ribbon:
    ![Smooth Lined Scatter](15.png)

3. In the resulting dialog box, repeatedly press the **Remove** button to remove all series. This is so we start with a blank slate:

    ![Clean slate](16.png)

4. Next click the **Add** button.
5. In the "Name" text box enter ```=Processed!H1```.
6. In the "X values" text box enter ```=Processed!EventDate```.
7. In the "Y values" text box enter ```=Processed!SunriseTime```.
8. Click the **Add** button again to add a second series.
5. In the "Name" text box enter ```=Processed!I1```.
6. In the "X values" text box enter ```=Processed!EventDate```.
7. In the "Y values" text box enter ```=Processed!SunsetTime```.

   Now attempt to add a series, using the same procedure, for day length. (Hint: Look at the screen grab below):

    ![Series](17.png)

8. Click the **OK** button to close the dialog box.

After resizing, your plot will look similar to this one:

![Series](18.png)

From this plot we can tell with some confidence that our data is accurate.
All three curves - like so many natural phemomena - follow an approximate sinusoid. 
The day length peaks during the summer months.
And reaches a minimum in the depths of winter.
The curve is smooth with no gaps and no outliers.

The final step, to make this graph really useful for publication or for later reference
is to add a title and ```x``` and ```y``` axis labels. In case you've forgotten how,
[here are some instructions](http://office.microsoft.com/en-au/excel-help/add-or-remove-titles-in-a-chart-HP001234162.aspx).

![Final chart](19.png)

<a class="next-link" href="{{ site.baseurl }}/module-7/">Go to the next module</a>
