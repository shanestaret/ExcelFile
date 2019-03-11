package com.company;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.util.IOUtils;

import java.io.InputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class Main {

    public static void main(String[] args) throws IOException {
        // below is to remove a row and shift rows up if duplicate rows are found (based on what first cell of each row contains); prints removed rows (based on index starting at 1 not 0)
        String fileName = "/Example/Excel/File/Example.xlsx";
        InputStream input = new FileInputStream(fileName);

        XSSFWorkbook wb = new XSSFWorkbook(input);
        XSSFSheet sheet = wb.getSheetAt(0);

        ArrayList<Integer> removedRows = new ArrayList<>();

        for(int i = 1; i < sheet.getLastRowNum() - 1; i++) {
            if (!removedRows.contains(i)) {
                XSSFRow rowInitial = sheet.getRow(i);
                XSSFCell cellInitial = rowInitial.getCell(0);
                String cellContentInitial = cellInitial.getRichStringCellValue().getString();

                for (int j = i + 1; j < sheet.getLastRowNum(); j++) {
                    if(!removedRows.contains(j)) {
                        XSSFRow rowCompare = sheet.getRow(j);
                        XSSFCell cellCompare = rowCompare.getCell(0);
                        String cellContentCompare = cellCompare.getRichStringCellValue().getString();

                        if (cellContentInitial.equals(cellContentCompare)) {
                            sheet.removeRow(rowCompare);
                            removedRows.add(j);
                            //sheet.shiftRows(j + 1, sheet.getLastRowNum(), -1); //uncomment this if there are no images (images would cause this to be wack otherwise)
                            FileOutputStream fileOut = null;
                            fileOut = new FileOutputStream("/Example/Excel/File/Example.xlsx");
                            wb.write(fileOut);
                            fileOut.close();
                        }
                    }
                }
            }
        }
        for(Integer rowNumber : removedRows) {
            System.out.print((rowNumber + 1) + " ");
        }

        // below is used to create an Excel spreadsheet and upload images from the computer to specific cells within that spreadsheet
        try {

            Workbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet("ExampleSheetName");
            for(int i = 1; i <= 660; i++) {
                //FileInputStream obtains input bytes from the image file
                InputStream inputStream = new FileInputStream("/Example/Logo/Path/logo " + i + ".png");
                //Get the contents of an InputStream as a byte[].
                byte[] bytes = IOUtils.toByteArray(inputStream);
                //Adds a picture to the workbook
                int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                //close the input stream
                inputStream.close();
                //Returns an object that handles instantiating concrete classes
                CreationHelper helper = wb.getCreationHelper();
                //Creates the top-level drawing patriarch.
                Drawing drawing = sheet.createDrawingPatriarch();

                //Create an anchor that is attached to the worksheet
                ClientAnchor anchor = helper.createClientAnchor();

                //create an anchor with upper left cell _and_ bottom right cell
                anchor.setCol1(0); //Column A
                anchor.setRow1(i - 1); //Row of image
                anchor.setCol2(1); //Column B
                anchor.setRow2(i); //Row below image

                //Creates a picture
                Picture pict = drawing.createPicture(anchor, pictureIdx);

                //Reset the image to the original size
                //pict.resize(); //don't do that. Let the anchor resize the image!

                //Create the Cell Ax and depending on what image is next
                Cell cell = sheet.createRow(0).createCell(i - 1);

                //set width to n character widths = count characters * 256
                int widthUnits = 20*256;
                sheet.setColumnWidth(1, widthUnits);

                //set height to n points in twips = n * 20
                short heightUnits = 60*20;
                cell.getRow().setHeight(heightUnits);

                //Write to Excel file
                FileOutputStream fileOut = null;
                fileOut = new FileOutputStream("/Example/Excel/File/Example.xlsx");
                wb.write(fileOut);
                fileOut.close();
            }

        }
        catch (IOException ioex) {
        }
    }
}
