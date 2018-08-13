import java.io.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GaitData {

    public void writeData() {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Ripon");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell (0);
            cell.setCellValue(3.5);
            FileOutputStream out = new FileOutputStream(new File("C:/Users/Protokinetics/Desktop/Colin/ExcelDemo.xlsx"));
            workbook.write(out);
            out.close();
        }
        catch(Exception e){
            System.out.println(e);
        }

        System.out.println("Excel file outputted");
        }
    }
