import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.sql.SQLOutput;
import java.util.Arrays;


public class PKMASParser {
    public static void main(String[] args) {

        PKMASParser edraw = new PKMASParser();

        String inputFileName = "C:/Users/Protokinetics/Desktop/Colin/backups/PKMAS.xlsx";
        String outputFileName = "C:/Users/Protokinetics/Desktop/Colin/backups/FINAL.xlsx";
        String location = "C:/Users/Protokinetics/Desktop/Colin";

        // Reading
        try (InputStream in = new FileInputStream(inputFileName)) {
            Workbook inWorkbook = WorkbookFactory.create(in);
            Sheet inSheet = inWorkbook.getSheetAt(0);

            DataFormatter formatter = new DataFormatter();

            Row inRow;
            Row row1, binaryOrFollowupRow;

            Cell inCell, reference, memo;
            Cell memoCheck, referenceCheck;

            // Count number of rows
            int inTotalRows = inSheet.getPhysicalNumberOfRows();
            int inTotalCols = 0;
            int inTmp = 0;

            // BLOCK: Counting number of rows/cells in input file
            for (int i = 0; i < 10 || i < inTotalRows; i++) {
                inRow = inSheet.getRow(i);
                if (inRow != null) {
                    inTmp = inSheet.getRow(i).getPhysicalNumberOfCells();
                    if (inTmp > inTotalCols) inTotalCols = inTmp;
                }
            }
            System.out.println("Excel file read");
            // End of reading block

            // BLOCK 1: calculates mean for each column

            Double BSMeanArray[] = new Double[44];
            Double BFMeanArray[] = new Double[44];
            Double FFSMeanArray[] = new Double[44];
            Double FFFMeanArray[] = new Double[44];
            Double SFSMeanArray[] = new Double[44];
            Double SFFMeanArray[] = new Double[44];

            for (int j = 5; j < 49; j++) {
                Double BSarray[] = new Double[inTotalRows];
                Double BFarray[] = new Double[inTotalRows];
                Double FFSarray[] = new Double[inTotalRows];
                Double FFFarray[] = new Double[inTotalRows];
                Double SFSarray[] = new Double[inTotalRows];
                Double SFFarray[] = new Double[inTotalRows];
                for (int i = 2; i < inTotalRows; i++) {
                    row1 = inSheet.getRow(i);
                    inCell = row1.getCell(j);

                    memoCheck = row1.getCell(4);
                    String memoCheckText = formatter.formatCellValue(memoCheck).toLowerCase().replaceAll("\\s", "");
                    referenceCheck = row1.getCell(3);
                    String referenceCheckText = formatter.formatCellValue(referenceCheck).toLowerCase().replaceAll("\\s", "");
                    if (memoCheckText.contains("baseline") && (referenceCheckText.equals("selfpace") || referenceCheckText.equals("self-pace"))) {
                        if (inCell != null) {
                            String text = formatter.formatCellValue(inCell);
                            Double textToDouble = Double.parseDouble(text);
                            BSarray[i] = textToDouble;
                        } else {
                            continue;
                        }
                    } else if (memoCheckText.contains("baseline") && (referenceCheckText.equals("fastpace") || referenceCheckText.equals("fast-pace"))) {
                        if (inCell != null) {
                            String text = formatter.formatCellValue(inCell);
                            Double textToDouble = Double.parseDouble(text);
                            BFarray[i] = textToDouble;
                        } else {
                            continue;
                        }
                    } else if ((memoCheckText.equals("follow-up") || memoCheckText.equals("followup") || memoCheckText.equals("1yearfollow-up") || memoCheckText.equals("1yearfollowup") || memoCheckText.equals("followup")) && (referenceCheckText.equals("selfpace") || referenceCheckText.equals("self-pace"))) {
                        if (inCell != null) {
                            String text = formatter.formatCellValue(inCell);
                            Double textToDouble = Double.parseDouble(text);
                            FFSarray[i] = textToDouble;
                        } else {
                            continue;
                        }
                    } else if ((memoCheckText.equals("follow-up") || memoCheckText.equals("followup") || memoCheckText.equals("1yearfollow-up") || memoCheckText.equals("1yearfollowup") || memoCheckText.equals("followup")) && (referenceCheckText.equals("fastpace") || referenceCheckText.equals("fast-pace"))) {
                        if (inCell != null) {
                            String text = formatter.formatCellValue(inCell);
                            Double textToDouble = Double.parseDouble(text);
                            FFFarray[i] = textToDouble;
                        } else {
                            continue;
                        }
                    }else if ((memoCheckText.equals("2yearfollow-up") || memoCheckText.equals("2yearfollowup")) && (referenceCheckText.equals("selfpace") || referenceCheckText.equals("self-pace"))) {
                        if (inCell != null) {
                            String text = formatter.formatCellValue(inCell);
                            Double textToDouble = Double.parseDouble(text);
                            SFSarray[i] = textToDouble;
                        } else {
                            continue;
                        }
                    }else if ((memoCheckText.equals("2yearfollow-up") || memoCheckText.equals("2yearfollowup")) && (referenceCheckText.equals("fastpace") || referenceCheckText.equals("fast-pace"))) {
                        if (inCell != null) {
                            String text = formatter.formatCellValue(inCell);
                            Double textToDouble = Double.parseDouble(text);
                            SFFarray[i] = textToDouble;
                        } else {
                            continue;
                        }
                    }
                }

                // calculating mean for column
                BSMeanArray[j - 5] = edraw.sum(BSarray) / edraw.numOfElements(BSarray);
                BFMeanArray[j - 5] = edraw.sum(BFarray) / edraw.numOfElements(BFarray);
                FFSMeanArray[j - 5] = edraw.sum(FFSarray) / edraw.numOfElements(FFSarray);
                FFFMeanArray[j - 5] = edraw.sum(FFFarray) / edraw.numOfElements(FFFarray);
                SFSMeanArray[j - 5] = edraw.sum(SFSarray) / edraw.numOfElements(SFSarray);
                SFFMeanArray[j - 5] = edraw.sum(SFFarray) / edraw.numOfElements(SFFarray);
                // test print array --TEST--
//                System.out.println(Arrays.toString(FFSarray));
//                System.out.println(edraw.sum(FFSarray));
//                System.out.println(edraw.numOfElements(FFSarray));
//                System.out.println(Arrays.toString(FFSMeanArray));
            }

            // BLOCK 2: check if baseline value or follow-up
            binaryOrFollowupRow = inSheet.getRow(2);
            reference = binaryOrFollowupRow.getCell(3);
            String referenceText = formatter.formatCellValue(reference);
            memo = binaryOrFollowupRow.getCell(4);
            String memoText = formatter.formatCellValue(memo);


            // BLOCK 3: output data
            InputStream outRead = new FileInputStream(outputFileName);
            Workbook outWorkbook = WorkbookFactory.create(outRead);
            Sheet outSheet1 = outWorkbook.getSheetAt(0);
            Sheet outSheet2 = outWorkbook.getSheetAt(1);

            Row outRow1, outRow2, row2, row3;
            Cell outCell;
            int outTotalRows1 = outSheet1.getPhysicalNumberOfRows();
            int outTotalRows2 = outSheet2.getPhysicalNumberOfRows();
            int outTotalCols1 = 0;
            int outTotalCols2 = 0;
            int outTmp1 = 0;
            int outTmp2 = 0;

            // BLOCK 4: Counting number of rows/cells in output file
            // THIS BLOCK IS REFERENCING THE ORIGINAL ROW COUNTING FROM INPUT FILE
            for (int i = 0; i < 10 || i < outTotalRows1; i++) {
                outRow1 = outSheet1.getRow(i);
                if (outRow1 != null) {
                    outTmp1 = outSheet1.getRow(i).getPhysicalNumberOfCells();
                    if (outTmp1 > outTotalCols1) outTotalCols1 = outTmp1;
                }
            }

            for (int i = 0; i < 10 || i < outTotalRows2; i++) {
                outRow2 = outSheet2.getRow(i);
                if (outRow2 != null) {
                    outTmp2 = outSheet2.getRow(i).getPhysicalNumberOfCells();
                    if (outTmp2 > outTotalCols2) outTotalCols2 = outTmp2;
                }
            }

            // BLOCK 5: Gets the arrays and prints to excel sheet
            for (int r = 4; r < outTotalRows1; r++) {
                row2 = outSheet1.getRow(r);
                row3 = outSheet2.getRow(r);
                if (row2.getCell(3) == null) {
                    for (int t = 3; t < 47; t++) {
                        outCell = row2.createCell(t);
                        outCell.setCellValue(BSMeanArray[t - 3]);
                    }
                    for (int u = 50; u < 94; u++) {
                        outCell = row2.createCell(u); // maybe this line wrong?
                        outCell.setCellValue(BFMeanArray[u - 50]);
                    }
                    for (int v = 3; v < 47; v++){
                        outCell = row3.createCell(v);
                        outCell.setCellValue(FFSMeanArray[v-3]);
                    }
                    for(int w = 50; w < 94; w++){
                        outCell = row3.createCell(w);
                        outCell.setCellValue(FFFMeanArray[w-50]);
                    }
                    for(int w = 97; w < 141; w++){
                        outCell = row3.createCell(w);
                        outCell.setCellValue(SFSMeanArray[w-97]);
                    }
                    for(int w = 144; w < 188; w++){
                        outCell = row3.createCell(w);
                        outCell.setCellValue(SFFMeanArray[w-144]);
                    }
                    break;

                }else{
                    continue;
                }
            }

            // BLOCK 6: PRINT TO EXCEL
            System.out.println("foo");
            FileOutputStream outWrite = new FileOutputStream(new File(outputFileName));
            outWorkbook.write(outWrite);
            outWrite.close();


        } catch (Exception e) {
            System.out.println(e);
            e.printStackTrace();
        }

// FUNCTIONS
    }

    private int numOfElements(Double[] array) {
        int count = 0;
        for (int i = 0; i < array.length; i++) {
            if (array[i] != null) {
                count++;
            }
        }
        return count;
    }


    private double sum(Double[] array) {
        double sum = 0;
        for (int i = 0; i < array.length; i++) {
            if (array[i] != null) {
                double number = array[i];
                sum += number;
            }
        }
        return sum;
    }

    private boolean isBaseline(String memoText) {
        String string = "baseline";
        if (memoText.toLowerCase().replaceAll("\\s", "").equals(string)) {
            return true;
        } else {
            return false;
        }
    }

    private boolean isFollowUp(String memoText) {
        String string = "follow-up";
        String string2 = "1yearfollowup";
        if (memoText.toLowerCase().replaceAll("\\s", "").equals(string) || memoText.toLowerCase().replaceAll("\\s", "").equals(string2)) {
            return true;
        } else {
            return false;
        }
    }

    private boolean isSecondFollowUp(String memoText) {
        if (memoText.toLowerCase().replaceAll("\\s", "").equals("2yearfollowup")) {
            return true;
        } else {
            System.out.println(memoText.toLowerCase());
            return false;
        }
    }

    private boolean isSelfPace(String referenceText) {
        String string = "selfpace";
        if (referenceText.toLowerCase().replaceAll("\\s", "").equals(string)) {
            return true;
        } else {
            return false;
        }
    }

    private boolean isFastPace(String referenceText) {
        String string = "fastpace";
        if (referenceText.toLowerCase().replaceAll("\\s", "").equals(string)) {
            return true;
        } else {
            System.out.println(referenceText.toLowerCase().replaceAll("\\s", ""));
            return false;
        }
    }
}
