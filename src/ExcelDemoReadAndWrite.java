import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.LinkedList;


public class ExcelDemoReadAndWrite {
    public static void main (String[]args){
        String fileName = "C:/Users/Protokinetics/Desktop/Colin/PKMAS.xlsx";
        String location = "C:/Users/Protokinetics/Desktop/Colin";

        // Reading
        try (InputStream in = new FileInputStream(fileName)){
            Workbook workbook = WorkbookFactory.create(in);
            Sheet inSheet = workbook.getSheetAt(0);

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            DataFormatter formatter = new DataFormatter();

            Row inRow;

            Row row, nextRow, prevRow;
            Cell patient, nextPatient, prevPatient;

            Cell inCell;

            // Count number of rows
            int rows = inSheet.getPhysicalNumberOfRows();
            int cols = 0;
            int tmp = 0;

            // Counting number of rows/cells
            for(int i = 0; i < 10 || i < rows; i++) {
                inRow = inSheet.getRow(i);
                if(inRow != null) {
                    tmp = inSheet.getRow(i).getPhysicalNumberOfCells();
                    if(tmp > cols) cols = tmp;
                }
            }
            System.out.println("Excel file read");

            Double array [] = new Double[rows];

            // -------------------------------------------------------
            Sheet PKMAS = workbook.getSheetAt(0);
//                for (Row row : PKMAS){

            for (int i = 2; i < rows; i++){
                row = inSheet.getRow(i);
                inCell = row.getCell(5);
                String text = formatter.formatCellValue(inCell);
                Double textToDouble = Double.parseDouble(text);
                array[i] = textToDouble;
            }
            // August 13
            for (int z = 0; z < array.length; z++){
                double number = array[z];
                double sum = 0;
            }
            double mean = sum/2;
            System.out.println(Arrays.toString(array));

            // -------------------------------------------------------

//            for(int r = 0; r < rows; r++) {
//                inRow = inSheet.getRow(r);
//                if(inRow != null) {
//                    for(int c = 0; c < cols; c++) {
//                        inCell = inRow.getCell((short)c);
//                        if(inCell != null) {
//                            // Writing execution block here
//
////                                Sheet outSheet = workbook.createSheet("Ripon");
////                                Row outRow = outSheet.createRow(0);
////                                Cell cell = outRow.createCell (0);
////                                cell.setCellValue(3.5);
//                                FileOutputStream out = new FileOutputStream(new File("C:/Users/Protokinetics/Desktop/Colin/ExcelDemo.xlsx"));
//                                workbook.write(out);
//                                out.close();
//
//                            System.out.println("Excel file outputted");
//                            break;
//                        }
//                    }
//                }
//            }
//            in.close();

        }catch(Exception e){
            System.out.println(e);
        }


    }
}
