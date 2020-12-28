# Project-Alva

## Background
A database matching project I did as a consultant for the Swedish Agency for Economic and Regional Growth, which was later dubbed Project Alva.
Originally a project that should have been done purely in Excel, a Python program has proven to be a better solution. The 1.1 is an earlier version with more redundant bits of code and no user interface. Offers more flexibility to programmers and testers. The final version offers a rough user interface for non-programmer users.

## Function

### Broadly
Takes in two Excel files and compares the selected properties in each row in one file to every row in the other file. Outputs an Excel file containing rows in the format (match value, row from first file, best matching row from second file).

### Algorithm
The program assumes that every row in the second file matches equally well with any row in the first file. In code this means that rather than just seeing the first file as a list of rows, like: [row 1, row_2, ... ,final row], the program takes this format and creates associations to the second file, making the format: [[row 1, match value 1, list of best matching rows in the other file 1, [row 2, match value 2, list of best matching rows in the other file 2], ... , [final row/match value/row match list]].

This means that every row in the first file now has an association to rows in the second file via the [row 1st file, match value, [list of best matching rows in 2nd file]] format. The initial assumption is that [list of best matching rows in 2nd file] contains every row in the 2nd file, meaning that every row in the second file matches equally well to a given row in the first file. These assumptions are then altered by comparing the row properties in a serial fashion, where only the best matches are kept each time. Ideally, this continues until the second list contains just a single, best-matching row.

First, the program matches "broad" properties like organisation numbers and financiers. The run sequence then determines what "fine" properties to search and in what order. Since only the best matches are kept after each parameter comparison, the matching order matters.

The program then changes this data back into a format that can be written in an Excel file and outputs this file in the same folder as it is run in.

## Future Work

### Input file cleanup
Excel input file cleanup does a lot of improve computation speed. In the project 25000 rows were compared with 7000 rows, meaning (25000*7000*=175 million) computations for a single property. For 5-6 properties it took 5-10 minutes until the program finished running. Every row did not correspond to a unique errand in the system - if there were several actors involved in an errand, they each get their own row in the system. If any redundant rows could be cleaned up without too much information loss, computation would be more efficient.

### User interface
The program was run directly from code using a Python interprefer and the user interface is more of a last-minute attempt to make the file more user-friendly, since this not a priority or even the intent of the project. This user interface can be more intuitive and less hard-coded. Ability to run custom sequences would be beneficial. This would require a function that translate run sequence code into matching commands. 

### Custom run sequences
This would enable manual run sequence input by reading run sequence code and outputs matching commands by making calls to the MatchBoxedLists function to use with specific conditions.

### Reduce amount of global variables
A lot of variables are declared in Section 2. Ideally, each variable should be declared in the function that uses them (except the main function) instead. See if this can be attended to.
