import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Date;

public class WriteExcel {

    public static void writeObject(XSSFSheet sheet, int row, int col, Object data) {

        try {
            Cell cell = null;
            XSSFRow sheetRow = sheet.getRow(row);
            if (sheetRow == null) {
                sheetRow = sheet.createRow(row);
            }
            cell = sheetRow.getCell(col);
            if (cell == null) {
                cell = sheetRow.createCell(col);
            }
            Object value = data;
            if (value instanceof String) {
                String s = (String) value;
                cell.setCellValue(s);
            } else if (value instanceof Double) {
                Double d = (Double) value;
                cell.setCellValue(d);
            } else if (value instanceof Integer) {
                Integer i = (Integer) value;
                cell.setCellValue(i);
            } else if (value instanceof Date) {
                Date date = (Date) value;
                cell.setCellValue(date);
            }

        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
    }

    public static void writeArrayList(XSSFSheet sheet, int startRow, int endRow, int startCol, int endCol, ArrayList data) {

        try {
            int index = 0;
            for (int row = startRow; row < endRow; row++) {
                for (int col = startCol; col < endCol; col++) {
                    Cell cell = null;
                    XSSFRow sheetRow = sheet.getRow(row);
                    if (sheetRow == null) {
                        sheetRow = sheet.createRow(row);
                    }
                    cell = sheetRow.getCell(col);
                    if (cell == null) {
                        cell = sheetRow.createCell(col);
                    }
                    Object value = data.get(index);
                    if (value instanceof String) {
                        String s = (String) value;
                        cell.setCellValue(s);
                    } else if (value instanceof Double) {
                        Double d = (Double) value;
                        cell.setCellValue(d);
                    } else if (value instanceof Integer) {
                        Integer i = (Integer) value;
                        cell.setCellValue(i);
                    } else if (value instanceof Date) {
                        Date date = (Date) value;
                        cell.setCellValue(date);
                    }
                    index++;
                }
            }

        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
    }

}
