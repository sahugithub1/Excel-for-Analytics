![excel](https://user-images.githubusercontent.com/91025550/148353577-6f740229-c993-4595-875f-7dcfd9a77a20.jpg)
# Excel-for-Analytics
## Tool : Excel

This project is designed from Matt Brattin's project tutorial https://youtu.be/45_yTM1HfTc
## Problem statement 
The board is asking to see how volume looked in Q2. I got some data (attached), but didn’t have a chance to pull anything together and was hoping you could take a stab at it.

I think they just want to see Q2 2021 volume by region and wanted to know if everything was looking good. I think this file has what you need. I don’t remember all the region codes – I know NAM ends in 1, EMEA ends in 3 and APAC and LATAM are 2 and 4, but I don’t remember which is which. I do know LATAM has the lowest volume so just go ahead and assign that to which ever comes out lowest.

## Shortcuts formula
I used these Alt shortcuts for productive work during this project.
1. Move tab right   ctrl+Page down
2. Move tab left    ctrl+page up
3. Move to end of dataset   ctrl+end
4. Move to top of sheet (A1)   ctrl+ Home
5. Move right across full cell   ctrl+ right arrow
6. Move down across full cell   ctrl+ down arrow
7. Select right across full cells   ctrl+shift+right arrow
8. Select down to end of cells   ctrl+shift+end

## Data cleaning

First of all save file with different name for source security.
copy each sheet tab and rename it ending with ( OG ) stands for original.
For rename tab click  Alt+H+O+R
Then hide it.

cleaning starts now 

Ctrl+A two times for full cells
for wrap/unwrap text alt+H+W
auto fit cell with Alt+H+O+I
then create table by ctrl+T
Rename table with Alt+J+T+A

## Filling blank cells
ctrl+spacebar for whole column selection
for selection of blank within range click Alt+H+F+D+S+K
then = and cell no which we want reference example =A2
for fill all left blanks click ctrl+enter
allcolumn value are now filled accordingly upper cell.

For grouping column value
select column by ctrl+space
Then text to column   Alt+A+E+F

## Reformatting dates and values from text
Reformat this ctrl + spacebar for selecting 
Alt+H+O+I, Alt+H+K, Alt+H+9

## Removing Duplicate
Alt+A+M

repeat all above process for both sheets.

## Checking string lengths and matching using RIGHT,MID,LEN functions

now used Xlookup, index match, and vlookup function for Region Id 

made table Start, end date and name of quarter example Q1 2021
Used Vlookup using approximate match for range created by us.

## Pivot table creation
for creating pivot table Alt+N+V
then created pivot chart
started variance analysis and basic forcasting.

## Donut Charts and Bar charts creation

## Key findings write-up ( Conclusion after Data cleaning, Data wrangling, Data Exploration and Analysis and creating Dashboard. )
Q2 YoY growth slowed from Q1 growth of 4% down to 2.7% or ~13k in volume primarily driven by:
(7k) volume or 55% of the decline from loss of two customers in LATAM driving overall growth for region from 9% down im Q1 to flat in Q2 YoY
NAM client onboarding in Q2 2020 anniversaried in Q2 2021, slowing perceived growth and amplifying Q1 growth by ~5k units or 1%
 
 ![Excel-for-Analytics](https://user-images.githubusercontent.com/91025550/137530352-0759a593-313c-4b83-b951-8380f0df86d6.png)
