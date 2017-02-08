/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.tutorial.excelreadwrite;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import javafx.scene.paint.Color;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
// up to 7 spaces
/**
 *
 * @author Andrew Hyun
 */
public class excelFunctions {
    
    private static XSSFWorkbook workbook;
    private static XSSFColor userDefinedColor;
    private static XSSFCellStyle mark;
    private static XSSFSheet sheet;
    private static boolean isCellMarked;
        
    excelFunctions(XSSFWorkbook workbook, int sheet){
        this.workbook = workbook;
        this.sheet = workbook.getSheetAt(sheet);
         
        //Define the color of mark -- This will be called in future functions
        mark = workbook.createCellStyle();
        mark.setFillForegroundColor(new XSSFColor(java.awt.Color.BLACK));
        mark.setFillPattern(CellStyle.SOLID_FOREGROUND);
    }
    
    /*
    ** Function #1: "convertColor"
    ** Obtains the user's defined foreground colors.
    ** Converts all of the colors to
    ** -------------User Defined Color-------------
    **
    ** NOTE: This funciton should be called FIRST to DEFINE userDefinedColor.
    */
    public void convertColor(int r, int g, int b, int numColors){   
        //Get the userDefinedColor and set the style
        userDefinedColor = new XSSFColor(new java.awt.Color(r, g, b));
        XSSFCellStyle userDefinedCS = workbook.createCellStyle();
        userDefinedCS.setFillForegroundColor(userDefinedColor);
        userDefinedCS.setFillPattern(CellStyle.SOLID_FOREGROUND);
        
        //Create an arrayList and add foreground colors that will be converted and then remove them
        List<XSSFColor> listOfColors = new ArrayList();
        for(int i = 0; i < numColors; ++i){
            try{
                //First row of excel document will be reserved for obtaining the colors of the foreground used
                listOfColors.add(sheet.getRow(0).getCell(i).getCellStyle().getFillForegroundXSSFColor());
                sheet.getRow(0).getCell(i).setCellStyle(null);
            }catch(NullPointerException ex){
                throw new NullPointerException("Either incorrect # colors entered OR colors NOT SET.");
            }
        }
        
        //Set-up rowIterator and get Row
        Iterator<Row> rowIterator = sheet.iterator();
        while(rowIterator.hasNext()){
            Row row = rowIterator.next();
            
            //Set-up cellIterator and get Cell
            Iterator<Cell> cellIterator = row.cellIterator();
            while(cellIterator.hasNext()){
                Cell cell = cellIterator.next();
                
                //Null-Check for Cell
                if(cell != null){
                    //Get the Cell Style, Null-Check for Cell Style
                    XSSFCellStyle currCellStyle = (XSSFCellStyle)cell.getCellStyle();
                    if(currCellStyle != null){
                        //Get the fillForeground color
                        XSSFColor fgColor = currCellStyle.getFillForegroundXSSFColor();
                        //cycle through ArrayList and compare if any of the colors listed matches
                        for(XSSFColor col : listOfColors){
                            if(col.equals(fgColor)){
                                cell.setCellStyle(userDefinedCS);
                            }
                        }
                    }
                }
                              
            }
        }
    }
    
    /*
    ** Function #2: "markHorizontal"
    ** Marks two cells that are
    ** --------------User Defined Unit--------------
    ** away HORIZONTALLY.    
    **
    ** NOTE: This function can ONLY BE CALLED AFTER "convertColor" has been 
    ** called EVERY SINGLE TIME. This is because the function will only read
    ** the userDefinedColor
    */
    public void markHorizontal(int spacesApart){
        //Set-up rowIterator and get Row
        Iterator<Row> rowIterator = sheet.iterator();
        while(rowIterator.hasNext()){
            Row row = rowIterator.next();
            
            //Set-up cellIterator and get Cell
            Iterator<Cell> cellIterator = row.cellIterator();
            while(cellIterator.hasNext()){
                Cell cell = cellIterator.next();
                
                //Obtains the Cell Style
                XSSFCellStyle cellStyle = (XSSFCellStyle)cell.getCellStyle();
                //Checks to see if the Cell Style is null; if null, go to next cell
                if(cellStyle != null){
                    //Checks to see what color is the cell's color
                    XSSFColor cellColor = cellStyle.getFillForegroundXSSFColor();
                    //Checks to see if color is null; if not compare to accept only editted or userDefined cells
                    if(cellColor != null){
                        //Checks if current cell is userDefined or editted
                        //If it is not, then go to the next cell
                        if(cellColor.equals(mark.getFillForegroundXSSFColor()) || cellColor.equals(userDefinedColor)){

                            //Set boolean isCellMarked to false before proceeding
                            isCellMarked = false;

                            //Define Cell to be (spacesApart+1) away
                            //So if x = current cell then the cell that is 5 spacesApart =
                            // [x][][][][][][x]
                            Cell cellMark = sheet.getRow(cell.getRowIndex()).getCell(cell.getColumnIndex() + spacesApart + 1);

                            //Checks to see if cell is null; if present, get its Cell Style
                            if(cellMark != null){
                                XSSFCellStyle cellMarkStyle = (XSSFCellStyle)cellMark.getCellStyle();

                                //Checks to see if the style is null; if present, get its color
                                if(cellMarkStyle != null){
                                    XSSFColor cellMarkColor = cellMarkStyle.getFillForegroundXSSFColor();

                                    //Checks to see if the color is null; if present, compare colors
                                    if(cellMarkColor != null){
                                        if(cellMarkColor.equals(userDefinedColor)){
                                            isCellMarked = true;
                                        }
                                    }
                                }
                            }

                            /*
                            ** CHECK#1: 'isCellMarked'
                            ** If isCellMarked is marked true, start iterating through the
                            ** cells in between and check if null or not userDefinedStyle
                            */
                            if(isCellMarked == true){
                                for(int i = 1; i <= spacesApart; ++i){
                                    Cell isNull = sheet.getRow(cell.getRowIndex()).getCell(cell.getColumnIndex()+i);

                                    //Checks to see if the cell is null; if color is present, set isCellMarked to false
                                    if(isNull != null){
                                        XSSFCellStyle cellCheckIfNullCellStyle = (XSSFCellStyle)isNull.getCellStyle();
                                        if(cellCheckIfNullCellStyle != null){
                                            XSSFColor cellCheckIfNullColor = cellCheckIfNullCellStyle.getFillForegroundXSSFColor();
                                            if(cellCheckIfNullColor != null){
                                                if(cellCheckIfNullColor.equals(userDefinedColor)){
                                                    isCellMarked = false;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            /*
                            ** CHECK#2: 'isCellMarked2'
                            ** If isCellMarked remains as true, set the two cell's style
                            */
                            if(isCellMarked == true){
                                cell.setCellStyle(mark);
                                cellMark.setCellStyle(mark);
                            }
                        }
                }
            }
        }
    }
}
