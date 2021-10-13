import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class mergeSheel  {

    mergeSheel()  throws IOException {

        //////// fichier 2 ////////////
        FileInputStream file2 = new FileInputStream(new File("D:\\excel\\1.xlsx"));
        //Get the workbook instance for XLS file
        XSSFWorkbook workbook2 = new XSSFWorkbook(file2);
        //Get first sheet from the workbook
        int sheetNumber = workbook2.getNumberOfSheets();





        String parcour="D:\\excel\\resultat.xlsx";
        FileOutputStream fileOut = new FileOutputStream(parcour);
        XSSFWorkbook worksheet = new XSSFWorkbook();

        String safeName = WorkbookUtil.createSafeSheetName("tout");
        XSSFSheet sheet = worksheet.createSheet(safeName);

        File fichier = new File(parcour);
        fichier.delete();
       int rownum=0;
        int rownum2=0;
        int rownum3=0;

        XSSFRow  rowEcrire = null;
        for (int i = 0;i< sheetNumber; i++) {
            XSSFSheet sheet2 = workbook2.getSheetAt(i);
            for (Row rowLire : sheet2) {
                if (isRowEmpty(rowLire)) {
                    continue;
                }
                rownum+=rowLire.getRowNum();
                System.out.println(rowLire.getRowNum()+" " +rownum);

                rowEcrire = sheet.createRow(rownum);
                XSSFCell cell =  rowEcrire.createCell(0);

                Cell cellLire = rowLire.getCell(0);
                cellLire.setCellType(Cell.CELL_TYPE_STRING);


                System.out.println(rownum);



                cell.setCellValue(cellLire.getStringCellValue());

            }



        }
        worksheet.write(fileOut);
        fileOut.close();
        fileOut.flush();

    }


    public static void main(String[] args) throws IOException {
        mergeSheel copm=new  mergeSheel();
    }


    private static boolean isRowEmpty(Row row) {
        boolean isEmpty = true;

        if (row != null) {
            for (Cell cell : row) {
                cell.setCellType(Cell.CELL_TYPE_STRING);

                if (cell.getStringCellValue().trim().length() > 0) {
                    isEmpty = false;
                    break;
                }
            }
        }

        return isEmpty;
    }
}
