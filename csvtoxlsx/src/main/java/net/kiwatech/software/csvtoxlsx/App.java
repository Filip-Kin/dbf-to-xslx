package net.kiwatech.software.csvtoxlsx;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ss.util.WorkbookUtil.createSafeSheetName;

public class App {
    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook();
        XSSFSheet sheet = (XSSFSheet) wb.createSheet(createSafeSheetName("Data"));

        if (readCSV(args[0], sheet) == 0) {
            System.out.println("Done! No data");
            return; // If theres only a header row then don't do anything
        }

        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();
        int firstCol = sheet.getRow(0).getFirstCellNum();
        int lastCol = sheet.getRow(0).getLastCellNum();

        CellReference topLeft = new CellReference(firstRow, firstCol);
        CellReference botRight = new CellReference(lastRow, lastCol - 1);
        AreaReference aref = wb.getCreationHelper().createAreaReference(topLeft, botRight);
        XSSFTable table = sheet.createTable(aref);

        // Repair the table
        for (int i = 0; i < table.getCTTable().getTableColumns().getCount(); i++) {
            table.getCTTable().getTableColumns().getTableColumnArray(i).setId(i+1);
        }
        table.setName("table");
        table.setDisplayName("Table");
        table.getCTTable().addNewTableStyleInfo();
        table.getCTTable().getTableStyleInfo().setName("TableStyleMedium2");

        FileOutputStream fileOut = new FileOutputStream(args[1]);
        wb.write(fileOut);
        fileOut.close();
        System.out.println("Done!");
    }

    public static int readCSV(String file, XSSFSheet sheet) throws IOException {
        int rowNum = 0;
        List<String> colNames = null;
        try (InputStream in = new FileInputStream(file);) {
            CSV csv = new CSV(true, ',', in);
            if (csv.hasNext()) {
                colNames = new ArrayList<String>(csv.next());
                Row row = sheet.createRow(0);
                for (int i = 0; i < colNames.size(); i++) {
                    String name = colNames.get(i);
                    row.createCell(i).setCellValue(name);
                }
            }

            boolean firstRow = true;
            while (csv.hasNext()) {
                List<String> fields = csv.next();
                if (firstRow) {
                    firstRow = false;
                    continue; // Skip the first line cause it's always blank (assuming this is from the script i wrote)
                }
                rowNum++;
                Row row = sheet.createRow(rowNum);
                for (int i = 0; i < fields.size(); i++) {
                    try {
                        double value = Double.parseDouble(fields.get(i));
                        row.createCell(i).setCellValue(value);
                    } catch (NumberFormatException ex) {
                        String value = fields.get(i);
                        try {
                            row.createCell(i).setCellValue(value);
                        } catch (IllegalArgumentException err) {
                            // Cell value too long, delete entire row
                            sheet.removeRow(row);
                        }
                    }
                }
            }
        }
        return rowNum;
    }
}
