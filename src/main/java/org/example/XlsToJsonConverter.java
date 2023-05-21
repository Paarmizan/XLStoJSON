package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

public class XlsToJsonConverter {
    private ObjectMapper mapper = new ObjectMapper();
    public JSONArray excelToJson(File inputFile, File readingParameters) throws IOException {

        // Json параметры
        Parameters parameters = mapper.readValue(readingParameters, Parameters.class);
        int processedSheet = parameters.getProcessedSheet();

        JSONObject excelData = new JSONObject();

        FileInputStream inputStream = null;
        Workbook workbook = null;

        try {

            inputStream = new FileInputStream(inputFile);

            String fileName = inputFile.getName().toLowerCase();
            if (fileName.endsWith(".xls") || fileName.endsWith(".xlsx")) {

                if (fileName.endsWith(".xls")) {
                    workbook = new HSSFWorkbook(inputStream);
                } else {
                    workbook = new XSSFWorkbook(inputStream);
                }


                Sheet sheet = workbook.getSheetAt(processedSheet - 1);
                String sheetName = sheet.getSheetName();


                int startRow = parameters.getStartRow() ;
                if (startRow > sheet.getLastRowNum()) {
                    throw new FaledTableInitialization("Начальная строка за пределами таблицы!");
                }
                int[] processedColumns = parameters.getProcessedColumns();
                for (int col : processedColumns) {
                    if (col > sheet.getRow(startRow).getLastCellNum()) {
                        throw new FaledTableInitialization("Выбран несуществующий стобец!");
                    }
                }
                int amountRecords = parameters.getAmountRecords();
                if (startRow - 1 + amountRecords > sheet.getLastRowNum() || amountRecords == -1) {
                    amountRecords = sheet.getLastRowNum() + 1;
                }




                JSONArray sheetData = new JSONArray();

                for (int i = startRow - 1; i < startRow - 1 + amountRecords; i++){
                    Row row = sheet.getRow(i);
                    JSONArray rowData = new JSONArray();
                    if (row == null) {
                        for (int j : processedColumns) {
                            String headerName = "Col" + j;
                            rowData.put(new JSONObject().put(headerName, ""));
                        }
                    } else {
                        for (int j : processedColumns) {
                            Cell cell = row.getCell(j - 1);
                            String headerName = "Col" + j;
                            if (cell != null) {
                                switch (cell.getCellType()) {
                                    case FORMULA:
                                        rowData.put(new JSONObject().put(headerName, cell.getCellFormula()));
                                        break;
                                    case BOOLEAN:
                                        rowData.put(new JSONObject().put(headerName, cell.getBooleanCellValue()));
                                        break;
                                    case NUMERIC:
                                        rowData.put(new JSONObject().put(headerName, cell.getNumericCellValue()));
                                        break;
                                    case BLANK:
                                        rowData.put(new JSONObject().put(headerName, ""));
                                        break;
                                    default:
                                        rowData.put(new JSONObject().put(headerName, cell.getStringCellValue()));
                                        break;
                                }
                            } else {
                                rowData.put(rowData);
                            }
                        }
                    }

                    sheetData.put(rowData);
                }

                excelData.append(sheetName, sheetData);
                System.out.println("Successful conversion!");
                return sheetData;

            } else {
                throw new IllegalArgumentException("Неподдерживаемый тип файла!");
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return null;
    }

}

class Parameters {
    private int startRow = 1;
    private int[] processedColumns = {1};
    private int amountRecords = -1;
    private int processedSheet = 1;
    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int[] getProcessedColumns() { return processedColumns; }

    public void setProcessedColumns(int[] processedColumns) { this.processedColumns = processedColumns; }

    public int getAmountRecords() {
        return amountRecords;
    }

    public void setAmountRecords(int amountRecords) {
        this.amountRecords = amountRecords;
    }
    public int getProcessedSheet() { return processedSheet; }

    public void setProcessedSheet(int processedSheet) { this.processedSheet = processedSheet; }
}
class FaledTableInitialization extends Exception {
    public FaledTableInitialization(String errorMessage) {
        super(errorMessage);
    }
}