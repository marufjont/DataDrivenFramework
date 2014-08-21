package com;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * Created by Maruf on 8/19/2014.
 */
public class ExcelMethods {
    public static String [][] readXL(String sPath, String iSheet) throws Exception{

        String [][] xData;


        File myxl = new File(sPath); // first create a file object providing the path to that file
        //Create the input stream from the xlsx/xls file
        FileInputStream fis = new FileInputStream(myxl);

        //Create Workbook instance for xlsx/xls file input stream
        Workbook myWB = null;
        if(sPath.toLowerCase().endsWith("xlsx")){
            myWB = new XSSFWorkbook(fis);
        }else if(sPath.toLowerCase().endsWith("xls")){
            myWB = new HSSFWorkbook(fis);
        }
        XSSFSheet mySheet = (XSSFSheet) myWB.getSheet(iSheet); // create object of a  specific sheet from the Excel file
        myWB.getSheet(iSheet);
        int xRows=mySheet.getLastRowNum()+1; // number of all rows
        int xCols=mySheet.getRow(0).getLastCellNum(); // number of all columns

        xData = new String [xRows][xCols];

        for (int i =0; i < xRows; i++){
            XSSFRow row =mySheet.getRow(i);
            for (int j=0;j<xCols;j++){
                XSSFCell cell = row.getCell(j);
                String value= cellToString(cell);

                xData[i][j]=value;
            }
        }
        return xData;

    }

    public static String cellToString(XSSFCell cell){
        //This function will convert an object of type excel to a string value
        int type =cell.getCellType();
        Object result;
        switch (type){
            case XSSFCell.CELL_TYPE_NUMERIC:
                //0
                result = cell.getNumericCellValue();

                break;
            case XSSFCell.CELL_TYPE_STRING:
                //1
                result = cell.getStringCellValue();

                break;
            case XSSFCell.CELL_TYPE_FORMULA:
                //2
                throw new RuntimeException("We can't evaluate formulas in Java");
            case XSSFCell.CELL_TYPE_BLANK:
                //3
                result = "%";
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                //4
                result = cell.getBooleanCellValue();
                break;

            case XSSFCell.CELL_TYPE_ERROR:
                //5
                throw new RuntimeException("This cell has an error");
            default:

                throw new RuntimeException("We  don't support this cell type " + type);
        }
        return result.toString();
    }
    //Method to write XL
    public static void writeExcel(String sPath, String iSheet, String xData[][]) throws Exception{

        File outFile = new File(sPath);
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet osheet=wb.createSheet(iSheet);
        int xR_TS = xData.length;
        int xC_TS = xData[0].length;

        for (int myrow = 0;myrow<xR_TS;myrow++){
            XSSFRow row = osheet.createRow(myrow);
            for (int mycol=0;mycol<xC_TS;mycol++){
                XSSFCell cell = row.createCell(mycol);
                cell.setCellType(	Cell.CELL_TYPE_STRING);
                cell.setCellValue(xData[myrow][mycol]);
            }
            FileOutputStream fOut = new FileOutputStream(outFile);
            wb.write(fOut);
            fOut.flush();
            fOut.close();
        }
    }
}
