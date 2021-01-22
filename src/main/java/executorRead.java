import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

public class executorRead {

    public static String XLSX_FULL_NAME = "./src/main/java/excel/java-excel-data.xlsx";

    public static void main(String[] args) {
        try {
            String sheetName = "Output";
            String cells = "";
            int row = 0;
            int col = 0;

            FileInputStream file = new FileInputStream(XLSX_FULL_NAME);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheet(sheetName);

            /* Nombre */
            cells = "C4";
            row = Utils.getDigit(cells) - 1;
            col = Utils.toNumber(Utils.getLetter(cells)) - 1;
            ReadExcel.readObject(sheet, row, col);

            /* Apellido */
            cells = "C5";
            row = Utils.getDigit(cells) - 1;
            col = Utils.toNumber(Utils.getLetter(cells)) - 1;
            ReadExcel.readObject(sheet, row, col);

            /* Edad */
            cells = "C6";
            row = Utils.getDigit(cells) - 1;
            col = Utils.toNumber(Utils.getLetter(cells)) - 1;
            ReadExcel.readObject(sheet, row, col);

            /* Notas */
            cells = "B9:C11";
            String[] cells_split = cells.split(":");
            String startCell = cells_split[0];
            String endCell = cells_split[1];
            int startRow = Utils.getDigit(startCell) - 1;
            int endRow = Utils.getDigit(endCell);
            int startCol = Utils.toNumber(Utils.getLetter(startCell)) - 1;
            int endCol = Utils.toNumber(Utils.getLetter(endCell));
            ReadExcel.readArrayList(sheet, startRow, endRow, startCol, endCol);

        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
    }
}
