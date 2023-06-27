package excel;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;

public class CompareTwoExcelFiles {

    public static void main(String[] args) {

        String user = System.getProperty("user.home")+"/desktop/testing files";
        String filePath1 = user+"/testing_excel_file.xlsx";
        String filePath2 = user+"/testing_excel_file_modified.xlsx";
        DataFormatter dataFormatter = new DataFormatter();

        LocalDateTime date = LocalDateTime.now();
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("dd-MM-yyyy HH-mm-ss");
        String formattedDate = date.format(dateTimeFormatter);



        try {

            Workbook workbook1 = new XSSFWorkbook(new FileInputStream(filePath1));
            Workbook workbook2 = new XSSFWorkbook(new FileInputStream(filePath2));

            Iterator<Sheet> sheetIterator1 = workbook1.sheetIterator();
            Iterator<Sheet> sheetIterator2 = workbook2.sheetIterator();

            while (sheetIterator1.hasNext() && sheetIterator2.hasNext()){
                Sheet sheet1 = sheetIterator1.next();
                Sheet sheet2 = sheetIterator2.next();

                Iterator<Row> rowIterator1 = sheet1.rowIterator();
                Iterator<Row> rowIterator2 = sheet2.rowIterator();

                while (rowIterator1.hasNext() && rowIterator2.hasNext()){
                    Row row1 = rowIterator1.next();
                    Row row2 = rowIterator2.next();

                    Iterator<Cell> cellIterator1 = row1.cellIterator();
                    Iterator<Cell> cellIterator2 = row2.cellIterator();

                    while (cellIterator1.hasNext() && cellIterator2.hasNext()){
                        Cell cell1 = cellIterator1.next();
                        Cell cell2 = cellIterator2.next();

                        String cellValue1 = dataFormatter.formatCellValue(cell1);
                        String cellValue2 = dataFormatter.formatCellValue(cell2);

                        if(!cellValue1.equals(cellValue2)){

                            CellStyle style = cell1.getSheet().getWorkbook().createCellStyle();
                            style.setFillForegroundColor(IndexedColors.RED.getIndex());
                            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            cell1.setCellStyle(style);

                        }
                    }
                }

            }

            FileOutputStream fileOutputStream = new FileOutputStream("comparison_output "+formattedDate+".xlsx");
            workbook1.write(fileOutputStream);
            workbook1.close();
            workbook2.close();
            fileOutputStream.close();


            System.out.println("comparison completed");


        }catch (FileNotFoundException e){
            System.out.println("check your file path");
        }catch (IOException e){
            System.out.println("IOException occurred");
        }





    }
}
