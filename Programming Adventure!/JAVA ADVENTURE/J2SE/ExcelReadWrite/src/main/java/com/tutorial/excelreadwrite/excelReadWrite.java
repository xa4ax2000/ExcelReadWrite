/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.tutorial.excelreadwrite;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Andrew Hyun
 */
public class excelReadWrite {
    
    private static XSSFWorkbook workbook;
    
    public static void main(String[] args){
/*************************************INPUT*************************************/ 
        System.out.println(new File(".").getAbsolutePath());
        String PATH = ("src\\main\\resources\\" + "betaTest.xlsx");
        try{
        FileInputStream inputStream = new FileInputStream(new File(PATH));
        workbook = new XSSFWorkbook(inputStream);
        }
        catch(IOException ex){
            ex.printStackTrace();
        }
/*************************************INPUT*************************************/        
/***********************************FUNCTIONS***********************************/
        excelFunctions fcn = new excelFunctions(workbook, 0);
        
        //User-defined Color = LIGHT_GREEN (144, 238, 144)
        //Number of colors to convert = 2
        fcn.convertColor(144, 238, 144, 2); 
        fcn.markHorizontal(7);
        
/***********************************FUNCTIONS***********************************/
/************************************OUTPUT*************************************/
        String PATH2 = ("src\\main\\resources\\" + "betaTestResults.xlsx");
        try{
            FileOutputStream outputStream = new FileOutputStream(new File(PATH2));
            workbook.write(outputStream);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
/************************************OUTPUT*************************************/
    }  
}