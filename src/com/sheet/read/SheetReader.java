package com.sheet.read;
 
import java.io.*;
import java.sql.*;
import java.util.*;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
 

public class SheetReader {
 
    public static void main(String[] args) {
        String jdbcURL = "jdbc:mysql://localhost:3308/assignment";
        String username = "root";
        String password = "root";
 
        String excelFilePath = "C:\\Users\\Desktop\\kushal\\sheet.xlsx";
 
        int batchSize = 20;
 
        Connection connection = null;
 
        try {
            long start = System.currentTimeMillis();
             
            FileInputStream inputStream = new FileInputStream(excelFilePath);
 
            workbook workbook = new XSSFWorkbook(inputStream);
 
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = firstSheet.iterator();
 
            connection = DriverManager.getConnection(jdbcURL, username, password);
            connection.setAutoCommit(false);
  
            String sql = "INSERT INTO students (id,name,lastname,identifier1,memberid,dob,effectiveDate,endDate1,voidFlag,contractId,batchA,state1,group1,division1,function1,Scenario1) VALUES (?, ?, ?,?,?, ?, ?,?,?, ?, ?,?,?, ?, ?,?,)";
            PreparedStatement statement = connection.prepareStatement(sql);    
             
            int count = 0;
             
            rowIterator.next(); // skip the header row
             
            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell nextCell = cellIterator.next();

                    int columnIndex = nextCell.getColumnIndex();

                   
                        switch (columnIndex) {
                        case 0:
                        	long id=(long)nextCell.getNumericCellValue();
                        	statement.setLong(1, id);
                          break;
                        case 1:
                        	String name=nextCell.getStringCellValue();
                        	statement.setString(2, name);
                        case 2:
                        	String lastname=nextCell.getStringCellValue();
                        	statement.setString(3, lastname);
                        case 3:
                        	String identifier1= nextCell.getStringCellValue();
                        	statement.setString(4, identifier1);
                        case 4:
                        	String memberid=nextCell.getStringCellValue();
                        	statement.setString(5, memberid);
                        case 5:
                        	String dob=nextCell.getStringCellValue();
                        	statement.setString(6, dob);
                        case 6:
                          String effectiveDate=nextCell.getStringCellValue();
                        	statement.setNString(7, effectiveDate);
                        case 7:
                        	String endDate1=nextCell.getStringCellValue();
                        	statement.setString(8, endDate1);
                        case 8:
                        	String voidFlag=nextCell.getStringCellValue();
                        	statement.setString(9, voidFlag);
                        case 9:
                        	String contractId=nextCell.getStringCellValue();
                        	statement.setString(10, contractId);
                        case 10:
                        	String batchA=nextCell.getStringCellValue();
                        	statement.setString(11, batchA);
                        case 11:
                        	String state1=nextCell.getStringCellValue();
                        	statement.setString(12, state1);
                        case 12:
                        	String group1=nextCell.getStringCellValue();
                        	statement.setString(13, group1);
                        case 13:
                        	String division1=nextCell.getStringCellValue();
                        	statement.setString(14, division1);
                        case 14:
                        	String function1=nextCell.getStringCellValue();
                        	statement.setString(15, function1);
                        case 15:
                        	String scenario1=nextCell.getStringCellValue();
                        	statement.setString(16, scenario1);
                       
                    }
                }
                 
                statement.addBatch();
                 
                if (count % batchSize == 0) {
                    statement.executeBatch();
                }              
 
            }
 
            workbook.close();
             
            // execute the remaining queries
            statement.executeBatch();
  
            connection.commit();
            connection.close();
             
            long end = System.currentTimeMillis();
            System.out.printf("Import done in %d ms\n", (end - start));
             
        } catch (IOException ex1) {
            System.out.println("Error reading file");
            ex1.printStackTrace();
        } catch (SQLException ex2) {
            System.out.println("Database error");
            ex2.printStackTrace();
        }
 
    }
}