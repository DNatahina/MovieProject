# Movie Dashboard
## Table of Content
[Problem Statement](#problem-statement)
[Data Source](#Data-source)
[Tools](#tools)
[]
### Problem Statement
Netflix wants to better understand which movie they should produce next, including the most suitable actors and directors. We have a dataset containing movie budgets, box office performance, actors, directors, and genres. Your task is to build an Excel dashboard that provides insights into this dataset. The dashboard should help identify:

- The best-performing actors
- The top movies based on box office metrics
- Director performance
- Genre trends
- Seasonal patterns in movie performance
- Any additional insights that can guide future production decisions
  
The final dashboard should be clear, interactive, and visually compelling, enabling Netflix to make data-driven decisions.
### Data Source
Movie Data : The primary dataset used for this analysis is the "Movie Data Homework.xlsx" file, containing detailed information about each movie's performance (box office and budget), actors, directors and genres.
You can download the original datasource here: [Movie Dataset Exel file](https://github.com/user-attachments/files/27219630/Movies_Data_Homework.xlsx)
### Tools
1. Power Query - I used Power Query for Data Cleaning
2. Excel - I used Excel for Data Analysis
3. Pivot Tables - for Creating the dashboard and Visualizations
### Data Cleaning
- Data loading and inspection.
- Handling errors, missing values.
- Data cleaning and formatting. The excel file after the data cleaning & preparation process can be downloaded here -[Movies Dashboard](https://github.com/user-attachments/files/27368150/movie_data_dashboard.xlsx)
### Dashboard
<img width="705" height="441" alt="image" src="https://github.com/user-attachments/assets/f3bd3bdf-7356-4780-9c88-897d3884c13a" />

### M Code

```
let
    Источник = Excel.Workbook(File.Contents("C:\Users\darya\Downloads\Movies_Data_Homework.xlsx"), null, true),
    #"Movie Data_Sheet" = Источник{[Item="Movie Data",Kind="Sheet"]}[Data],
    #"Повышенные заголовки" = Table.PromoteHeaders(#"Movie Data_Sheet", [PromoteAllScalars=true]),
    #"Измененный тип" = Table.TransformColumnTypes(#"Повышенные заголовки",{{"Movie Title", type text}, {"Release Date", type date}, {"Wikipedia URL", type text}, {"Genre_First_ID", Int64.Type}, {"Genre_Second_ID", Int64.Type}, {"Director_First_ID", Int64.Type}, {"Cast_First_ID", Int64.Type}, {"Cast_Second_ID", Int64.Type}, {"Cast_Third_ID", Int64.Type}, {"Cast_Fourth_ID", Int64.Type}, {"Cast_Fifth_ID", Int64.Type}, {"Budget ($)", Int64.Type}, {"Box Office Revenue ($)", type number}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type any}}),
    #"Объединенные запросы" = Table.NestedJoin(#"Измененный тип", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Удаленные столбцы" = Table.RemoveColumns(#"Объединенные запросы",{"Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21"}),
    #"Развернутый элемент Genres" = Table.ExpandTableColumn(#"Удаленные столбцы", "Genres", {"Genre"}, {"Genres.Genre"}),
    #"Переупорядоченные столбцы" = Table.ReorderColumns(#"Развернутый элемент Genres",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Переименованные столбцы" = Table.RenameColumns(#"Переупорядоченные столбцы",{{"Genres.Genre", "Genre"}}),
    #"Объединенные запросы1" = Table.NestedJoin(#"Переименованные столбцы", {"Genre_Second_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Развернутый элемент Genres1" = Table.ExpandTableColumn(#"Объединенные запросы1", "Genres", {"Genre"}, {"Genres.Genre"}),
    #"Переупорядоченные столбцы1" = Table.ReorderColumns(#"Развернутый элемент Genres1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genres.Genre", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Переименованные столбцы1" = Table.RenameColumns(#"Переупорядоченные столбцы1",{{"Genres.Genre", "Genre_Second"}}),
    #"Объединенные запросы2" = Table.NestedJoin(#"Переименованные столбцы1", {"Director_First_ID"}, Directors, {"ID"}, "Directors", JoinKind.LeftOuter),
    #"Развернутый элемент Directors" = Table.ExpandTableColumn(#"Объединенные запросы2", "Directors", {"Director"}, {"Directors.Director"}),
    #"Переупорядоченные столбцы2" = Table.ReorderColumns(#"Развернутый элемент Directors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Переименованные столбцы2" = Table.RenameColumns(#"Переупорядоченные столбцы2",{{"Directors.Director", "Director"}}),
    #"Объединенные запросы3" = Table.NestedJoin(#"Переименованные столбцы2", {"Cast_First_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Развернутый элемент Actors" = Table.ExpandTableColumn(#"Объединенные запросы3", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Переименованные столбцы3" = Table.RenameColumns(#"Развернутый элемент Actors",{{"Actors.Actor", "Actor"}}),
    #"Переупорядоченные столбцы3" = Table.ReorderColumns(#"Переименованные столбцы3",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Объединенные запросы4" = Table.NestedJoin(#"Переупорядоченные столбцы3", {"Cast_Second_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Развернутый элемент Actors1" = Table.ExpandTableColumn(#"Объединенные запросы4", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Переименованные столбцы4" = Table.RenameColumns(#"Развернутый элемент Actors1",{{"Actors.Actor", "Actor_2"}}),
    #"Переупорядоченные столбцы4" = Table.ReorderColumns(#"Переименованные столбцы4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Actor_2", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Объединенные запросы5" = Table.NestedJoin(#"Переупорядоченные столбцы4", {"Cast_Third_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Развернутый элемент Actors2" = Table.ExpandTableColumn(#"Объединенные запросы5", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Переупорядоченные столбцы5" = Table.ReorderColumns(#"Развернутый элемент Actors2",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Actor_2", "Cast_Third_ID", "Actors.Actor", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Переименованные столбцы5" = Table.RenameColumns(#"Переупорядоченные столбцы5",{{"Actors.Actor", "Actor_3"}}),
    #"Объединенные запросы6" = Table.NestedJoin(#"Переименованные столбцы5", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Развернутый элемент Actors3" = Table.ExpandTableColumn(#"Объединенные запросы6", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Переупорядоченные столбцы6" = Table.ReorderColumns(#"Развернутый элемент Actors3",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Actor_2", "Cast_Third_ID", "Actor_3", "Cast_Fourth_ID", "Actors.Actor", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Переименованные столбцы6" = Table.RenameColumns(#"Переупорядоченные столбцы6",{{"Actors.Actor", "Actor_4"}}),
    #"Объединенные запросы7" = Table.NestedJoin(#"Переименованные столбцы6", {"Cast_Fifth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Развернутый элемент Actors4" = Table.ExpandTableColumn(#"Объединенные запросы7", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Переупорядоченные столбцы7" = Table.ReorderColumns(#"Развернутый элемент Actors4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Actor_2", "Cast_Third_ID", "Actor_3", "Cast_Fourth_ID", "Actor_4", "Cast_Fifth_ID", "Actors.Actor", "Budget ($)", "Box Office Revenue ($)"}),
    #"Переименованные столбцы7" = Table.RenameColumns(#"Переупорядоченные столбцы7",{{"Actors.Actor", "Actor_5"}}),
    #"Добавлен пользовательский объект" = Table.AddColumn(#"Переименованные столбцы7", "ROI", each ([#"Box Office Revenue ($)"]-[#"Budget ($)"])/[#"Budget ($)"]),
    #"Измененный тип1" = Table.TransformColumnTypes(#"Добавлен пользовательский объект",{{"ROI", Percentage.Type}})
in
    #"Измененный тип1"
```
### Recommendations
Top 5 genres are Action, Comedy, etc. I would recommend Netfliz to produce a movie with one of these genres as they brought in more in box office revenie based on the data from 2012 to 2016
<img width="284" height="86" alt="image" src="https://github.com/user-attachments/assets/5b77f0bc-6e20-427a-9918-42710a3a0311" />



