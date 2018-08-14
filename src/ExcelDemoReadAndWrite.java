import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Arrays;


public class ExcelDemoReadAndWrite {
    public static void main (String[]args){

        ExcelDemoReadAndWrite edraw = new ExcelDemoReadAndWrite();

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
            // End of reading block

            Double array [] = new Double[rows];

            // block that calculates mean for each column
            for (int j = 5; j < 43; j++) {
                System.out.println();
                for (int i = 2; i < 6; i++) {
                    row = inSheet.getRow(i);
                    inCell = row.getCell(j);
                    String text = formatter.formatCellValue(inCell);
                    System.out.println(text);
                    Double textToDouble = Double.parseDouble(text);
                    array[i] = textToDouble;
                    System.out.println(Arrays.toString(array));
                }
                // calculating mean for column
                System.out.println(edraw.sum(array)/edraw.numOfElements(array));
            }


            // -------------------------------------------------------
            // write to file
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
            // --------------------------------------------------------

        }catch(Exception e){
            System.out.println(e);
        }

// FUNCTIONS
    }

    private int numOfElements(Double[]array){
        int count = 0;
        for(int i = 0; i < array.length; i++){
            if(array[i] != null){
                count++;
            }
        }
        return count;
    }


    private double sum(Double[] array ){
        double sum = 0;
        for(int i = 0; i < array.length; i++){
            if(array[i] != null){
                double number = array[i];
                sum += number;
            }
        }
        return sum;
    }
}
