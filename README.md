
````markdown
# ExcelGenerator

This Java project creates an Excel (.xlsx) file with tabular user data, highlights errors, and applies formatting using the Apache POI library.

## Features

- Creates an Excel file named `users.xlsx` with sample data.
- Highlights the header row with a blue background.
- Sets fixed column widths.
- Highlights cells with errors:
  - Age less than 0 is highlighted with a red background.
  - Email missing the `@` symbol is highlighted with a yellow background.
  - Rows with status "Error" are highlighted entirely with a gray background.

## Requirements

- Java 11 or higher
- Apache POI library (included in project or via build tool)

## How to run

1. Compile the project:

   bash
   javac -cp ".;path/to/poi-library/*" task_1/ExcelGenerator.java
````

2. Run the program:

   ```bash
   java -cp ".;path/to/poi-library/*" task_1.ExcelGenerator
   ```

3. After running, check the generated file `users.xlsx` in the project directory.

## Notes

* Make sure to include Apache POI dependencies in your classpath.
* The program currently uses hardcoded data for demonstration purposes.
* The Excel file will be overwritten each time you run the program.

## Author

Kirill Reges
