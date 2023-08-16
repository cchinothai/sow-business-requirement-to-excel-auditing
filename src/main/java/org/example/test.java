//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//
//public class ExcelWriter {
//    public static void main(String[] args) {
//        String filePath = "path/to/existing/file.xlsx";
//
//        try (FileInputStream fis = new FileInputStream(filePath)) {
//            // Load the existing workbook
//            Workbook workbook = new XSSFWorkbook(fis);
//            Sheet sheet = workbook.getSheetAt(0); // Assuming you want to work with the first sheet
//
//            String desiredValue = "146"; // The value you are looking for in column A
//
//            // Find the row with the desired value in column A
//            int desiredRowIndex = -1;
//            for (Row row : sheet) {
//                Cell cell = row.getCell(0); // Assuming you are looking in column A (index 0)
//                if (cell != null && cell.getCellType() == CellType.STRING && desiredValue.equals(cell.getStringCellValue())) {
//                    desiredRowIndex = row.getRowNum();
//                    break;
//                }
//            }
//
//            if (desiredRowIndex != -1) {
//                // Create a new row at the found index
//                Row newRow = sheet.createRow(desiredRowIndex + 1); // Index + 1 to insert after the found row
//
//                // Example data to insert into specific cells
//                String data1 = "New Data";
//                int data2 = 42;
//
//                // Insert data into specific cells in the newly created row
//                Cell cellA = newRow.createCell(0); // Column A
//                cellA.setCellValue(data1);
//
//                Cell cellB = newRow.createCell(1); // Column B
//                cellB.setCellValue(data2);
//            } else {
//                System.out.println("Desired value not found in the specified column.");
//            }
//
//            // Save the changes back to the same file
//            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
//                workbook.write(fileOut);
//                System.out.println("Data inserted and changes saved successfully.");
//            }
//            workbook.close();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
//}