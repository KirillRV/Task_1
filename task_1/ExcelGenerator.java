package task_1;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

/**
 * The {@code ExcelGenerator} class creates an Excel (.xlsx) file
 * with sample tabular data, applies formatting and highlights errors
 * based on specific validation rules.
 *
 * <p>Highlighting rules:</p>
 * <ul>
 *     <li>Header row — light blue background.</li>
 *     <li>Age &lt; 0 — red background.</li>
 *     <li>Email without '@' — yellow background.</li>
 *     <li>Status "Error" — entire row with gray background.</li>
 * </ul>
 *
 * <p>Uses the Apache POI library for Excel file generation.</p>
 *
 * @author YourName
 */
public class ExcelGenerator {

    /**
     * The main entry point of the program.
     * Creates and saves an Excel file named "users.xlsx".
     *
     * @param args command-line arguments (not used)
     */
    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Users");

        List<String[]> data = Arrays.asList(
                new String[]{"Name", "Email", "Age", "Status"},
                new String[]{"Ivan", "ivan@example.com", "25", "OK"},
                new String[]{"Olga", "olga@example", "-1", "Error"},
                new String[]{"Alexey", "alexeyexample.com", "30", "OK"},
                new String[]{"Maria", "maria@example.com", "-5", "OK"},
                new String[]{"Sergey", "sergey@example.net", "40", "Error"}
        );

        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle redStyle = createColorStyle(workbook, IndexedColors.RED);
        CellStyle yellowStyle = createColorStyle(workbook, IndexedColors.YELLOW);
        CellStyle grayRowStyle = createRowStyle(workbook, IndexedColors.GREY_25_PERCENT);

        //loop that applies gray background to entire row when status is "Error"
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i);
            String[] rowData = data.get(i);
            boolean isErrorRow = i > 0 && "Error".equalsIgnoreCase(rowData[3]);

            for (int j = 0; j < rowData.length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(rowData[j]);

                if (i == 0) {
                    // Header style
                    cell.setCellStyle(headerStyle);
                } else {
                    if (isErrorRow) {
                        // Apply gray background to entire row when status is "Error"
                        cell.setCellStyle(grayRowStyle);
                    } else {
                        // Check email validity
                        if (j == 1 && !rowData[j].contains("@")) {
                            cell.setCellStyle(yellowStyle);
                        }
                        // Check age value
                        else if (j == 2) {
                            try {
                                int age = Integer.parseInt(rowData[j]);
                                if (age < 0) {
                                    cell.setCellStyle(redStyle);
                                }
                            } catch (NumberFormatException e) {
                                cell.setCellStyle(redStyle);
                            }
                        }
                    }
                }
            }
        }

        // Set fixed column widths
        sheet.setColumnWidth(0, 15 * 256); // Name
        sheet.setColumnWidth(1, 20 * 256); // Email
        sheet.setColumnWidth(2, 10 * 256); // Age
        sheet.setColumnWidth(3, 10 * 256); // Status

        // Save file
        try (FileOutputStream fileOut = new FileOutputStream("users.xlsx")) {
            workbook.write(fileOut);
            workbook.close();
            System.out.println("File users.xlsx has been created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Creates a style for the header row (light blue background, bold font).
     *
     * @param workbook the Excel workbook
     * @return the cell style for headers
     */
    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * Creates a cell style with a specified background color.
     *
     * @param workbook the Excel workbook
     * @param color the background color
     * @return the colored cell style
     */
    private static CellStyle createColorStyle(Workbook workbook, IndexedColors color) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    /**
     * Creates a row style with the specified background color (for error rows).
     *
     * @param workbook the Excel workbook
     * @param color the background color
     * @return the row style
     */
    private static CellStyle createRowStyle(Workbook workbook, IndexedColors color) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }
}
