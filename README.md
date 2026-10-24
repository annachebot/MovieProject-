# Movie Dashboard 
## Table of Content 
[Problem Statement](#problem-statement)
[Data Source](#data-source)
[Tools](#tools)
[]
[]
### Problem Statement
Netflix wants to better understand which movie they should produce next, including the most suitable actors and directors.

We have a dataset containing movie budgets, box office performance, actors, directors, and genres.

Your task is to build an Excel dashboard that provides insights into this dataset.

The dashboard should help identify:

- The best-performing actors

- The top movies based on box office metrics

- Director performance

- Genre trends

- Seasonal patterns in movie performance

- Any additional insights that can guide future production decisions
  
The dashboard should be clear, interactive, and visually compelling, enabling Netflix to make data-driven decisions.
### Data Source 
Movie Data : The primary dataset used for this analysis is the "Movie Data Homework.xlsx" file, containing detailed information about each movie's performance (box office and budget), actors, directors and genres. 
You can download the original datasource here: [Movies_Data_Homework.xlsx](https://github.com/user-attachments/files/33913241/Movies_Data_Homework.xlsx)

### Tools
1. Power Query - I used Power Query for Data Cleaning
2. Excel - I used Excel for Data Analysis
3. Pivot Tables - for Creating the dashboard and Visualizations
### Data Cleaning 
- Data loading and inspection.
- Handling errors, missing values.
- Data cleaning and formatting. The excel file after the data cleaning & preparation process
can be downloaded here -[Movies_Dashboard.xlsx](https://github.com/user-attachments/files/33913360/Movies_Data_Homework_Anna.xlsx)
### Dashboard 
<img width="771" height="630" alt="image" src="https://github.com/user-attachments/assets/3575e2e3-8f7f-40a5-8ea8-bce6d8c5d56f" />

### M Code

```
let
    Source = Excel.Workbook(File.Contents("C:\Users\andre\OneDrive\Рабочий стол\Anna\Data Analitics Course\Movies_Data_Homework.xlsx"), null, true),
    #"Movie Data_Sheet" = Source{[Item="Movie Data",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Movie Data_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Movie Title", type text}, {"Release Date", type date}, {"Wikipedia URL", type text}, {"Genre_First_ID", Int64.Type}, {"Genre_Second_ID", Int64.Type}, {"Director_First_ID", Int64.Type}, {"Cast_First_ID", Int64.Type}, {"Cast_Second_ID", Int64.Type}, {"Cast_Third_ID", Int64.Type}, {"Cast_Fourth_ID", Int64.Type}, {"Cast_Fifth_ID", Int64.Type}, {"Budget ($)", Int64.Type}, {"Box Office Revenue ($)", type number}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type any}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres" = Table.ExpandTableColumn(#"Merged Queries", "Genres", {"Genre"}, {"Genres.Genre"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Genres",{"Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries1" = Table.NestedJoin(#"Reordered Columns", {"Genre_Second_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres1" = Table.ExpandTableColumn(#"Merged Queries1", "Genres", {"Genre"}, {"Genres.Genre.1"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded Genres1",{{"Genres.Genre.1", "Genres second"}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Renamed Columns",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres second", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries2" = Table.NestedJoin(#"Reordered Columns1", {"Director_First_ID"}, Directors, {"ID"}, "Directors", JoinKind.LeftOuter),
    #"Expanded Directors" = Table.ExpandTableColumn(#"Merged Queries2", "Directors", {"Director"}, {"Directors.Director"}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Expanded Directors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries3" = Table.NestedJoin(#"Reordered Columns2", {"Cast_First_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors" = Table.ExpandTableColumn(#"Merged Queries3", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Reordered Columns3" = Table.ReorderColumns(#"Expanded Actors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries4" = Table.NestedJoin(#"Reordered Columns3", {"Cast_Second_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Reordered Columns4" = Table.ReorderColumns(#"Merged Queries4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Expanded Actors1" = Table.ExpandTableColumn(#"Reordered Columns4", "Actors", {"Actor"}, {"Actors.Actor.1"}),
    #"Merged Queries5" = Table.NestedJoin(#"Expanded Actors1", {"Cast_Third_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Reordered Columns5" = Table.ReorderColumns(#"Merged Queries5",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors.Actor.1", "Cast_Third_ID", "Actors", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Expanded Actors2" = Table.ExpandTableColumn(#"Reordered Columns5", "Actors", {"Actor"}, {"Actors.Actor.2"}),
    #"Merged Queries6" = Table.NestedJoin(#"Expanded Actors2", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Reordered Columns6" = Table.ReorderColumns(#"Merged Queries6",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors.Actor.1", "Cast_Third_ID", "Actors.Actor.2", "Cast_Fourth_ID", "Actors", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Expanded Actors3" = Table.ExpandTableColumn(#"Reordered Columns6", "Actors", {"Actor"}, {"Actors.Actor.3"}),
    #"Merged Queries7" = Table.NestedJoin(#"Expanded Actors3", {"Cast_Fifth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors4" = Table.ExpandTableColumn(#"Merged Queries7", "Actors", {"Actor"}, {"Actors.Actor.4"}),
    #"Reordered Columns7" = Table.ReorderColumns(#"Expanded Actors4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors.Actor.1", "Cast_Third_ID", "Actors.Actor.2", "Cast_Fourth_ID", "Actors.Actor.3", "Cast_Fifth_ID", "Actors.Actor.4", "Budget ($)", "Box Office Revenue ($)"}),
    #"Added Custom" = Table.AddColumn(#"Reordered Columns7", "ROE", each ([#"Box Office Revenue ($)"]-[#"Budget ($)"])/[#"Budget ($)"]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom",{{"ROE", Percentage.Type}})
in
    #"Changed Type1"
```
