package org.example;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.nio.file.Paths;
import java.io.IOException;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class PDFAudit {

    //store all of the business requirement IDs
    //read in a text file containing PCI's design document
    //iterate through all of the BR IDs and check to see if they occur within the design document
    public static void main(String[] args) {
        //store all business requirement IDs here
        /*
        - Put all id numbers in a text file separated by the space
        - Read in text file and split into words
        - prepend "BR." to each id
        - add id's into BRIDs
         */
        List<String> BRIDs = new ArrayList<>();

        String txtFilePath = "C:\\Users\\CChinoth\\ETRM-Auditing\\id-group-1.txt";
        String txtToString = fileContentToString(txtFilePath);
        String[] words = txtToString.split("\n");

        for(String word: words){

            BRIDs.add("BR."+word);
        }


        //filepath for PCI's design document
        String pdfFilePath = "C:\\Users\\CChinoth\\ETRM-Auditing\\PCI-Deployment-Design-SDGE ETRM vX (23_07_12).pdf";

        try{
            List<String> missingIDs = findMissingBRIDs(pdfFilePath, BRIDs);
            if(missingIDs.isEmpty()){
                System.out.println("All the business requirement IDs are found in the PCI design document");
            }
            else{
//                System.out.println("The following business requirement IDs are explicitly missing from the PCI design document:");
//                for(int i = 0; i < missingIDs.size(); i++){
//                    System.out.println(missingIDs.get(i));
//                }
                System.out.println("Number of BR IDs searched: "+ BRIDs.size());
                System.out.println("Number of explicitly missing BR IDs: "+ missingIDs.size());
            }
        } catch (IOException e){
            e.printStackTrace();
        }


    }

    public static List<String> findMissingBRIDs(String filePath, List<String> keywords) throws IOException{
        List<String> missingKeywords = new ArrayList<>();

        //File path to working excel book
        String excelFile = "C:\\Users\\CChinoth\\ETRM-Auditing\\ETRM BR Auditing.xlsx";



        //Read in pdf file
        try(PDDocument document = PDDocument.load(new File(filePath))) {
            //create instance of pdfTextStripper
            PDFTextStripper pdfts = new PDFTextStripper();

            //retrieve text content from PDF
            String text = pdfts.getText(document);

            //split the text into words
            String[] words = text.split("\\s+"); //split by all whitespace characters


            //store words containing "BR."
            List<String> foundBRIDs = new ArrayList<>();
            for(String word: words){
                if (word.contains("BR.")) { //BR.001
                    String formattedWord = word.substring(0,6);
                    foundBRIDs.add(formattedWord);
                }
                if (word.contains("BR-")){
                    System.out.println("HERE"+word);
                    if(word.length()>4){
                        String cut = word.substring(3,6);
                        String paste = "BR."+cut;

                        foundBRIDs.add(paste);
                    }
                }
            }

            for (String id : keywords) {
                boolean found = false;
                for (String foundID : foundBRIDs) {
                    if (id.contains(foundID)) {
                        found = true;
                        break;
                    }
                }
                if (!found) {
                    missingKeywords.add(id);
                    //EXCEL OPERATION: update cell to elevate check
                    updateWorkbook(excelFile,id,false);

                }
                else{
                    //EXCEL OPERATION: update cell to confirm check
                    updateWorkbook(excelFile,id,true);
                }

            }
            return missingKeywords;
        }
    }

    public static void updateWorkbook(String filePath, String brID, boolean found) throws IOException {
        try(FileInputStream fis = new FileInputStream(filePath)){
            Workbook etrmBRAuditing = new XSSFWorkbook(fis);
            Sheet excelSheet = etrmBRAuditing.getSheetAt(0);

            String id = brID.substring(3);
            //System.out.println(id); //001 -> 1
            //**TASK: Need to format to get rid of all zeroes
            String replace = id.replaceAll("^0+", "").replaceAll("\\D", "");
            int excelFormattedID = Integer.valueOf(replace);



            int findRowIndex = findExcelRow(excelFormattedID,excelSheet);

            if(findRowIndex != -1){
                Row row = excelSheet.getRow(findRowIndex);

                String data;
                if(found){data="Y";}
                else{ data = "?";}

                Cell cell = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cell.setCellValue(data);
            }
            else {
                System.out.println("Desired value not found in the specified column.");
            }

            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                etrmBRAuditing.write(fileOut);
                System.out.println("Data modified and changes saved successfully.");
            }
            etrmBRAuditing.close();
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static int findExcelRow(int id, Sheet sheet){
        System.out.println(id);
        for (Row row : sheet) {
            if(row.getRowNum()==0){continue;}
            Cell cell = row.getCell(0); // Assuming you are looking in column A (index 0)
            //System.out.println(cell.getNumericCellValue());
            if (cell != null && cell.getCellType() == CellType.NUMERIC && id == cell.getNumericCellValue()) {
                return row.getRowNum();
            }
        }
        return -1;
    }

    private static String fileContentToString(String filePath){
        String fileContent;
        try{
            byte[] bytes = Files.readAllBytes(Paths.get(filePath));
            fileContent = new String(bytes, StandardCharsets.UTF_8);
            return fileContent;
        } catch (IOException e){
            e.printStackTrace();
        }
        return "Unable to convert file content to string";
    }

}