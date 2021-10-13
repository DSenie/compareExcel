import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class compare {



    compare() throws IOException, InvalidFormatException {

        //////// fichier 1 ////////////

        FileInputStream file = new FileInputStream(new File("D:\\excel\\1.xlsx"));
        Workbook workbook = WorkbookFactory.create(file);

        Sheet sheet = workbook.getSheetAt(4);

      /*  //Get the workbook instance for XLS file
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        //Get first sheet from the workbook
        XSSFSheet sheet = workbook.getSheetAt(1);*/


        //////// fichier 2 ////////////
        FileInputStream file2 = new FileInputStream(new File("D:\\excel\\2.xlsx"));
        //Get the workbook instance for XLS file
        XSSFWorkbook workbook2 = new XSSFWorkbook(file2);
        //Get first sheet from the workbook
        XSSFSheet sheet2 = workbook2.getSheetAt(0);





        for (Row row : sheet) {

            for (Row row2 : sheet2) {


                if (isRowEmpty(row)) {
                    continue;
                }
                Cell cell = row.getCell(0);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                String matricule = cell.getStringCellValue();


                Cell cell2 = row2.getCell(0);
                cell2.setCellType(Cell.CELL_TYPE_STRING);
                String matricule2 = cell2.getStringCellValue();



                if(matricule.trim().equals(matricule2.trim())){

                 ////////recuperer les valeur

                    ////////nom
                    Cell cellNom = row2.getCell(2);
                    cellNom.setCellType(Cell.CELL_TYPE_STRING);
                    String Nom = cellNom.getStringCellValue();

                    ///////prenom
                    Cell cellPrenom = row2.getCell(3);
                    cellPrenom.setCellType(Cell.CELL_TYPE_STRING);
                    String Prenom = cellPrenom.getStringCellValue();





                    /////////ecrire

                    /////nom
                    Cell cellNomE= row.createCell( 3);
                    cellNomE.setCellValue(Nom);

                    ///Prenom
                    Cell cellPrenonE = row.createCell( 4);
                    cellPrenonE.setCellValue(Prenom);



                }
            }
        }


        file.close();

        FileOutputStream outputStream = new FileOutputStream("D:\\excel\\1.xlsx");
        workbook.write(outputStream);
        //workbook.close();
        outputStream.close();





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






    public static void main(String[] args) throws IOException, InvalidFormatException {
      compare copm=new  compare();
    }







}
