package org.example;

import org.json.JSONArray;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {

        try {
            String inputFilePath = "C://Users//k//Desktop//excel//table.xlsx";
            String readingParametersPath = "C://Users//k//Desktop//excel//readingParameters.json";
            String outputFolderPath = "C://Users//k//Desktop//excel//";
            File inputFile = new File(inputFilePath);
            File outputFolder = new File(outputFolderPath);
            File readingParameters = new File(readingParametersPath);

            XlsToJsonConverter xlsToJsonConverter = new XlsToJsonConverter();
            JSONArray result = xlsToJsonConverter.excelToJson(inputFile, readingParameters);

            String outputFileName;
            if (inputFile.getName().endsWith(".xls")) {
                outputFileName = inputFile.getName().replace(".xls", ".json");
            } else {
                outputFileName = inputFile.getName().replace(".xlsx", ".json");
            }

            File outputFile = new File(outputFolder, outputFileName);
            FileWriter fileWriter = new FileWriter(outputFile);
            fileWriter.write(result.toString());
            fileWriter.flush();

        } catch (IOException ex) {
            ex.printStackTrace();
        }


    }
}