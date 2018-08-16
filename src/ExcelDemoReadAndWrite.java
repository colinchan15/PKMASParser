import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Arrays;


public class ExcelDemoReadAndWrite {
    public static void main (String[]args){

        ExcelDemoReadAndWrite edraw = new ExcelDemoReadAndWrite();

        String inputFileName = "C:/Users/Protokinetics/Desktop/Colin/ExcelDemo.xlsx";
        String outputFileName = "C:/Users/Protokinetics/Desktop/Colin/GAIT_ANALYSIS_JAVA_TEST.xlsx";
        String location = "C:/Users/Protokinetics/Desktop/Colin";

        // Reading
        try (InputStream in = new FileInputStream(inputFileName)){
            Workbook inWorkbook = WorkbookFactory.create(in);
            Sheet inSheet = inWorkbook.getSheetAt(0);

            DataFormatter formatter = new DataFormatter();

            Row inRow;
            Row row1, binaryOrFollowupRow;

            Cell inCell, reference, memo;

            // Count number of rows
            int inTotalRows = inSheet.getPhysicalNumberOfRows();
            int inTotalCols = 0;
            int inTmp = 0;

            // BLOCK: Counting number of rows/cells in input file
            for(int i = 0; i < 10 || i < inTotalRows; i++) {
                inRow = inSheet.getRow(i);
                if(inRow != null) {
                    inTmp = inSheet.getRow(i).getPhysicalNumberOfCells();
                    if(inTmp > inTotalCols) inTotalCols = inTmp;
                }
            }
            System.out.println("Excel file read");
            // End of reading block

            // BLOCK: calculates mean for each column
            Double array [] = new Double[inTotalRows];
            Double meanArray [] = new Double[43];

            for (int j = 5; j < 48; j++) {
//                System.out.println();
                for (int i = 2; i < inTotalRows; i++) { // must change the 6 to be updated to max rows
                    row1 = inSheet.getRow(i);
                    inCell = row1.getCell(j);
                    String text = formatter.formatCellValue(inCell);
                    Double textToDouble = Double.parseDouble(text);
                    array[i] = textToDouble;
                }
                // calculating mean for column
                // test print array contents
//                System.out.println(edraw.sum(array)/edraw.numOfElements(array));
                meanArray[j-5] = edraw.sum(array)/edraw.numOfElements(array);
                // test print array
//                System.out.println(Arrays.toString(meanArray));
            }

            // BLOCK: check if baseline value or follow-up
            binaryOrFollowupRow = inSheet.getRow(2);
            reference = binaryOrFollowupRow.getCell(3);
            String referenceText = formatter.formatCellValue(reference);
            memo = binaryOrFollowupRow.getCell(4);
            String memoText = formatter.formatCellValue(memo);





            // BLOCK: output data
            InputStream outRead = new FileInputStream(outputFileName);
            Workbook outWorkbook = WorkbookFactory.create(outRead);
            Sheet outSheet = outWorkbook.getSheetAt(0);

            Row outRow, row2;
            Cell outCell;
            int outTotalRows = 0;
            int outTotalCols = 0;
            int outTmp = 0;

            // BLOCK: Counting number of rows/cells in output file
            // THIS BLOCK IS REFERENCING THE ORIGINAL ROW COUNTING FROM INPUT FILE
            for(int i = 0; i < 10 || i < outTotalRows; i++) {
                outRow = outSheet.getRow(i);
                if(outRow != null) {
                    outTmp = outSheet.getRow(i).getPhysicalNumberOfCells();
                    if(outTmp > outTotalCols) outTotalCols = outTmp;
                }
            }


            // check if baseline or follow-up first
            if (edraw.isBaseline(memoText) == true) {
                if (edraw.isSelfPace(referenceText) == true) {
                    for (int r = 4; r < outTotalRows; r++) {
                        row2 = outSheet.getRow(r);
                        if (row2.getCell(3) == null) { // if the velocity cell is == null, then set cell value to array established and then break
                            for (int t = 3; t < 26; t++) {
                                outCell = row2.createCell(t); // maybe this line wrong?
                                outCell.setCellValue(meanArray[t-3]);
                                FileOutputStream outWrite = new FileOutputStream(new File(outputFileName));
                                outWorkbook.write(outWrite);
                                outWrite.close();
                            }
                            break;
                        }
                    }
                }
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
            in.close();
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

    private boolean isBaseline(String memoText){
        String string = "baseline";
        if(memoText.toLowerCase().equals(string)){
            return true;
        }else {
            System.out.println(memoText.toLowerCase());
            return false;
        }
    }

    private boolean isSelfPace (String referenceText){
        String string = "selfpace";
        if(referenceText.toLowerCase().replaceAll("\\s", "").equals(string)){
            return true;
        }else{
            System.out.println(referenceText.toLowerCase().replaceAll("\\s",""));
            return false;
        }
    }
}
