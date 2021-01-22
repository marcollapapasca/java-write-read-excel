import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

public class executorWrite {

    public static String XLSX_FULL_NAME = "./src/main/java/excel/java-excel-data.xlsx";

    public static void main(String[] args) {

        try {
            String sheetName = "Input";
            String cells = "";
            int row = 0;
            int col = 0;
            Object data = "";

            FileInputStream file = new FileInputStream(XLSX_FULL_NAME);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheet(sheetName);

            /* Nombre */
            cells = "C4";
            data = "Marco";
            row = Utils.getDigit(cells) - 1;
            col = Utils.toNumber(Utils.getLetter(cells)) - 1;
            WriteExcel.writeObject(sheet, row, col, data);

            /* Apellido */
            cells = "C5";
            data = "Llapapasca";
            row = Utils.getDigit(cells) - 1;
            col = Utils.toNumber(Utils.getLetter(cells)) - 1;
            WriteExcel.writeObject(sheet, row, col, data);

            /* Edad */
            cells = "C6";
            data = 32;
            row = Utils.getDigit(cells) - 1;
            col = Utils.toNumber(Utils.getLetter(cells)) - 1;
            WriteExcel.writeObject(sheet, row, col, data);

            /* Notas */
            cells = "C9:E11";
            String[] cells_split = cells.split(":");
            String startCell = cells_split[0];
            String endCell = cells_split[1];
            int startRow = Utils.getDigit(startCell) - 1;
            int endRow = Utils.getDigit(endCell);
            int startCol = Utils.toNumber(Utils.getLetter(startCell)) - 1;
            int endCol = Utils.toNumber(Utils.getLetter(endCell));
            ArrayList arrayList = new ArrayList();
            arrayList.add(12);
            arrayList.add(12);
            arrayList.add(18);
            arrayList.add(9);
            arrayList.add(15);
            arrayList.add(12);
            arrayList.add(8);
            arrayList.add(12);
            arrayList.add(16);
            WriteExcel.writeArrayList(sheet, startRow, endRow, startCol, endCol, arrayList);

            /*-------------------------*/
            file.close();

            FileOutputStream outFile = new FileOutputStream(new File(XLSX_FULL_NAME));
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            workbook.write(outFile);
            outFile.close();

        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
    }

}
