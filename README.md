# CsvToKml
A CSV to KML generator intended for generating visualisation of radio networks

This tool takes as input a table of locations with the following format:

```
Name,Lat,Long,Group
```
Lat and Long are decimal latitue and decimal longitude coordinates.  The Goup field is used to group the locations by folder in google earth, and if the Group name also exists as a location then a path will be drawn between the two.

The format of the table can be either csv, or excel spreadsheet.  Be sure to name the header row as per the above (the order doesn't matteer).  The Group column can be blank, but the header must exist.

Optionally, a table of point to point links can also be included either as an additional csv file, or as an additional worksheet titles "Links".

The format of the links table is
```
Name,Location1,Location2
```
Where Name is the name of the link, Location1 is the location of one end of the link (must match an entry in the Locations table exactly) and Location2 is the other end.

Usage: 
```
csvtokml file1 [file2]
```
file1 can be either a csv list of locations, or an xml spreadsheet containing locations and links
file2 is an optional csv list of links

Two output files are generated, file1.kml and file1_nolabels.kml.

