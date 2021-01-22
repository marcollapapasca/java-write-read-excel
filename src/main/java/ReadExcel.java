import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;

public class ReadExcel {

    public static void readObject(XSSFSheet sheet, int row, int col) {
        try {
            Object data = "";
            Cell cell = sheet.getRow(row).getCell(col);
            switch (cell.getCellType()) {
                case STRING:
                    data = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    data = cell.getNumericCellValue();
                    break;
                case BOOLEAN:
                    data = cell.getBooleanCellValue();
                    break;
                case FORMULA:
                    switch (cell.getCachedFormulaResultType()) {
                        case NUMERIC:
                            data = cell.getNumericCellValue();
                            break;
                        case STRING:
                            data = cell.getRichStringCellValue();
                            break;
                    }
                    break;
                default:
                    data = cell.getCellFormula();
            }

            System.out.println(data);

        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
    }

    public static void readArrayList(XSSFSheet sheet, int startRow, int endRow, int startCol, int endCol) {

        try {
            ArrayList data = new ArrayList();
            for (int row = startRow; row < endRow; row++) {
                for (int col = startCol; col < endCol; col++) {
                    Cell cell = sheet.getRow(row).getCell(col);
                    switch (cell.getCellType()) {
                        case STRING:
                            data.add(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            data.add(cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            data.add(cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            switch (cell.getCachedFormulaResultType()) {
                                case NUMERIC:
                                    data.add(cell.getNumericCellValue());
                                    break;
                                case STRING:
                                    data.add(cell.getRichStringCellValue());
                                    break;
                            }
                            break;
                        default:
                            data.add("");
                    }
                }
            }
            System.out.println(data);

        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
    }

}
