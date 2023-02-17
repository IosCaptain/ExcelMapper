Excel Mapper to DBs leveraging Apache POI. 

1. In the main method we pass the file location to create a FileStream and the excel workbook from there. We then get the row and column count to use
in the loops. The first loop is for rows and the second one for the number of the cell in that row, aka, column.
2. Then we can specify the field of the entity we want to assign depending on the the column and store them in a list until we are ready to save into the DB
3. There is something different in the email fields to allow more than one email in the cell so we are splitting the string on "; "


There are a couple of cool thing here:
- 2 dimensions loops 
- some work with strings towards the bottom
- switch case block
- java streams to filter nulls out
- jpa repository use

<img src="https://img.shields.io/badge/Spring-6DB33F?style=for-the-badge&logo=spring&logoColor=white" /> <img src="https://img.shields.io/badge/apache_maven-C71A36?style=for-the-badge&logo=apachemaven&logoColor=white" />
