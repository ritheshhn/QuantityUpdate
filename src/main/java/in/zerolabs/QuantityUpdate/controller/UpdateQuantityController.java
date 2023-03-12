package in.zerolabs.QuantityUpdate.controller;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;

@RestController
@RequestMapping("ctrl")
public class UpdateQuantityController {

    @GetMapping("/fetch")
    public void fetchDetailsFromExcel() throws IOException {

        System.out.println("<<<<<  UploadExcelDetails  >>  fetchDetailsFromExcel()  >>>>>");

        String hyperFilePath = "T:\\Projects\\POCs\\QuantityUpdate\\QuantityUpdate\\src\\main\\resources\\US_Hypercare_Sheet.xlsx";
        File hyperFile = new File(hyperFilePath);
        FileInputStream hyperFIS = new FileInputStream(hyperFile);


        FileInputStream eodFIS = new FileInputStream(new File("T:\\Projects\\POCs\\QuantityUpdate\\QuantityUpdate\\src\\main\\resources\\EOD.xls"));

        XSSFWorkbook hyperCareworkbook = new XSSFWorkbook(hyperFIS);
        HSSFWorkbook eodworkbook = new HSSFWorkbook(eodFIS);

        XSSFSheet hyperCareWorksheet = hyperCareworkbook.getSheetAt(0);
        HSSFSheet eodWorksheet = eodworkbook.getSheetAt(0);

        System.out.println("<<<<< No of Owners present in Excel = "
                + String.valueOf((hyperCareWorksheet.getPhysicalNumberOfRows())) + "  >>>>>");

        for (int i = 1; i < hyperCareWorksheet.getPhysicalNumberOfRows(); i++) {
            System.out.println("i = "+i);

            XSSFRow row = hyperCareWorksheet.getRow(i);

            Cell cell = row.getCell(2);

            String s = cell.getStringCellValue();
            String hyperDelNo = s.substring(1);

            System.out.println(hyperDelNo);

            for (int j = 1; j < eodWorksheet.getPhysicalNumberOfRows()-1; j++) {
                HSSFRow eodRow = eodWorksheet.getRow(j);

                Cell eodCell = eodRow.getCell(1);

                String str = String.valueOf(eodCell.getNumericCellValue());
                String eodDelNo = str.substring(2, str.length()-2);



                if(hyperDelNo.equalsIgnoreCase(eodDelNo)){
                    System.out.println(hyperDelNo + "  "+eodDelNo);
                    System.out.println("Equalllllllllllllllll");

                    Cell orderQuantityCell = eodRow.getCell(10);
                    int orderQuantity = (int) orderQuantityCell.getNumericCellValue();

                    row.getCell(3).setCellValue(orderQuantity);

                }
            }

            /*for (int j = 0; j < 33; j++) {
                //System.out.println("j = "+j);
                Cell cell = row.getCell(j);

                System.out.println(cell.getStringCellValue());


            }*/
        }

        FileOutputStream os = new FileOutputStream(hyperFile);
        hyperCareworkbook.write(os);
        os.close();



    }
}
