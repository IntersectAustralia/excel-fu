---
layout: post
title: Where to next? 
---

### Evaluation Survey
<div class="note">
Please complete our <a href="http://svy.mk/18c8dHa">course evaluation survey</a>. Your feedback helps us to make <b>Excel Fu: Excel for Researchers</b> even better.
</div>
<!---
### Ask me anything / Show and Tell sessions (online)
<div class="note">
Join us online for an <b>Ask me anything / Show and Tell session</b>. These sessions are a great opportunity to ask any Excel questions your might have, or to show us what you've done with your data after completing <b>Excel Fu</b>:

<ul>
  <li> <a href="http://intersect-excel-fu-hangout2.eventbrite.com.au">3 October 2014</a> </li>
  <li> <a href="http://intersect-excel-fu-hangout3.eventbrite.com.au">17 October 2014</a> </li>

</ul>
</div>
--->
<!---
### Join the Mailing List
<div class="note">
  Find out about upcoming courses by <a href="http://bit.ly/1aZvRqw">signing up to Intersect's courses mailing list</a>.
</div>
--->

### More cool stuff

<div class="note">
  See below for ways you can do more with Excel. And some other handy tools and techniques you might like to
  explore...
</div>

## Microsoft Excel Analysis ToolPak

The [Microsoft Excel Analysis ToolPak](http://office.microsoft.com/en-au/excel-help/use-the-analysis-toolpak-to-perform-complex-data-analysis-HP010342762.aspx) is an addon you can load in Excel that allows you perform more complex analysis in Excel. 

The pack provides a collection of tools that support:

* correlation
* covariance
* Fourier analysis
* moving averages
* histograms
* random number generation
* sampling
* and many other advanced functions

Check out the link above for the full list of features.

## Data entry with Forms

If you are doing a lot of data entry in Excel, the worksheet format can be prone to errors. It can be worth _locking it down_, to only permit valid values to be entered. It might even be worth creating a data entry form, specifically laid out to permit efficient data entry.

Here are a couple of videos that might be helpful:

* [Excel Power Tips: 10 Ways to make Data Entry Faster and More Accurate](https://www.youtube.com/watch?v=6BErKQ29Jjc)
* [How to: Create a simple Userform in Excel](https://www.youtube.com/watch?v=5oXcct1mOUw)

There are many more resources on the internet.

## Don't repeat yourself!

If you catch yourself repeating the same action over and over, consider automating it. Excel provides **macros** for this purpose. Macros let you "record" commonly completed actions and "replay" that sequence of actions.

Check out [this video](http://office.microsoft.com/en-us/excel-help/video-work-with-macros-VA104034656.aspx?CTT=5&origin=HA104032083) for a quick introduction to macros.

Closely related is Visual Basic for Applications (VBA). This is a programming language based on [BASIC](http://en.wikipedia.org/wiki/BASIC). At the most basic level, VBA lets you write custom functions, which you can then use in your worksheets.

At a more advanced level you can use VBA to customise and automate just about everything in your workbook.

For a taste, check out some of these [tutorials on VBA](http://www.excel-easy.com/vba.html).

## Auditing 

We only scratched the surface of auditing your spreadsheet for errors.

Check out [this tutorial](http://www.excel-easy.com/examples/formula-auditing.html) on formula auditing. There are lots of cool features, such as **tracing precedents** that help to reveal what's going on in your spreadheet. And don't forget the **show formulas** option we used earlier.

These features are also great for working out what the heck is going on in a workbook someone else has emailed you.

## Clustering and Open Refine
Often data is noisy, messy and inconsistent and needs cleaning before it can be analysed. You may, for instance, have a worksheet with UFO sightings by country. The "country" column is unfortunately really inconsistent. Australia for instance is written in different ways:

* Australia
* AU
* australia

The United States is worse:

* United States
* US
* U.S.
* USA
* U.S.A.
* United States of America

These kinds of inconsistences play havoc when trying to aggregate your data.

While **filtering**, **sorting** and particularly **pivot tables** are wonderful Excel features that let you get a handle on your data, you might also like to take a look at [Open Refine](http://openrefine.org/) (formerly _Google Refine_). 

Open Refine offers pivot table style functionality through what it calls *facets*. But it also offers [cluster analysis](http://en.wikipedia.org/wiki/Cluster_analysis), which means it can be used to automatically detect and fix inconsistencies. It doesn't always work perfectly, but it can be a huge time saver.

Open Refine doesn't do everything Excel does and vice versa. Together they can be used to do some pretty powerful analysis.

Intersect offers an [introductory course on Open Refine](http://www.intersect.org.au/training) that you might like to attend.

## Databases and SQL

Excel workbooks used to be limited to 65,536 rows, now they support over 1 million. But if your data has lots of records, you might want to look at moving to a database.

If your data is relational, moving to a database will most likely simplify extracting the data you need. A key clue that your data is relational is that you have lots of duplicated data in your worksheet.

**Access** and **FileMaker Pro** are popular databases with point-and-click interfaces. If you favour the command line, **sqlite**, **MySQL**, **PostgreSQL** are powerful databases you can start using for free. 

Software Carpentry offers a great introduction to [databases and SQL](http://software-carpentry.org/v5/novice/sql/index.html) that uses **sqlite**.

## Programming and R

When doing advanced analysis, when working with huge amounts of data, or when wanting to completely automate analysis, it's hard to look past a programming language.

[R](http://www.r-project.org/) has gained popularity as a statistical computing language. It's well worth taking a look at if you want to do a lot of statistical analysis.

[Python](https://www.python.org/) is a popular language in the scientific community and is also worth exploring as more general programming language. Software Carpentry offer [an introduction](http://software-carpentry.org/v5/novice/python/index.html).

If you'd like to learn some key programming concepts in a graphical and friendly way, check out [Scratch](http://scratch.mit.edu/).

## Version Control

Backing up, taking copies, using file naming conventions, manual versioning, and protecting sheets all offer protection from data loss.

But nothing beats version control software for guarding against ever losing anything important. Popular options are [Git](http://git-scm.com/) and [Subversion](http://subversion.apache.org/). These allow you to create *repositories* for your files. After making changes to a file (say your Excel workbook) you check it into the repository. The repository keeps all the versions and allows you to extract the file as it stood at the time of any given check in.

You might like to explore [this Software Carpentry introduction](http://software-carpentry.org/v5/novice/git/index.html).


