Excel Mapper to DBs leveraging Apache POI. 

In the main method we pass the file location to create a FileStream and the excel workbook from there. We then get the row and column count to use
in the loops. The first loop is for rows and the second one for the number of the cell in that row, aka, column.
Then we can specify the field of the entity we want to assign depending on the the column and store them in a list until we are ready to save into the DB
There is something different in the email fields to allow more than one email in the cell so we are splitting the string on "; "


There are a couple of cool thing here:
- 2 dimensions loops 
- some work with strings towards the bottom
- switch case block
- java streams to filter nulls out
- jpa repository use
