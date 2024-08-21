package org.example;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;

@RestController
public class HomeController {

    @PostMapping("/upload")
    public ResponseEntity<InputStreamResource> uploadFile(@RequestParam("file") MultipartFile file) throws IOException {
        try {
            // Convert CSV to Excel
            ByteArrayInputStream in = new ByteArrayInputStream(file.getBytes());
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            convertCsvToExcel(in, out);

            // Return the Excel file as a response
            InputStreamResource resource = new InputStreamResource(new ByteArrayInputStream(out.toByteArray()));
            HttpHeaders headers = new HttpHeaders();
            headers.add("Content-Disposition", "attachment; filename=converted.xlsx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentLength(out.size())
                    .body(resource);
        } catch (IOException e) {
            e.printStackTrace();
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    private void convertCsvToExcel(ByteArrayInputStream csvInputStream, ByteArrayOutputStream excelOutputStream) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        try (CSVParser csvParser = new CSVParser(new InputStreamReader(csvInputStream), CSVFormat.DEFAULT.withFirstRecordAsHeader())) {
            int rowNum = 0;
            for (CSVRecord csvRecord : csvParser) {
                Row row = sheet.createRow(rowNum++);
                int cellNum = 0;
                for (String field : csvRecord) {
                    row.createCell(cellNum++).setCellValue(field);
                }
            }
        }

        workbook.write(excelOutputStream);
        workbook.close();
    }

}
