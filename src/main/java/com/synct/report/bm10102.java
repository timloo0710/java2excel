package com.synct.report;
import java.io.*;
import java.util.*;
import java.sql.*;
import java.lang.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;
import com.synct.util.*;
import com.codecharge.*;
import com.codecharge.components.*;
import com.codecharge.util.*;
import com.codecharge.events.*;
import com.codecharge.db.*;
import com.codecharge.validation.*;
//
public class bm10102 extends Ole2Adapter {

	private		int onepage_detail = 20000;     //
	private    	int dtl_start_row = 3;      //
	private    	int dtl_cols = 2;           //
	private    	String execlfilename = "bm10102.xls";  //

    public bm10102() {
        page_rows = 20000;     //
	}


    //
    private String separator;
    private String path;

    public void setReset(String emptyString){
      this.path="";
      this.path="";
   }

    public void setPath(String path){
         this.path=path;
    }

    public String getPath(){
       return this.path;
    }

  public POIFSFileSystem fs;
  public HSSFWorkbook wb;
	public HSSFSheet sheet;
  public HSSFSheet sheet1;
  public HSSFSheet sheet2;
  public HSSFSheet sheet3;
	public HSSFPrintSetup ps;

	HSSFCellStyle[][] header_style;   //
	HSSFCellStyle[][] body_style;     //

	String[] data;
	String[][] header_value;          //
	String[][] body_value;            //

	Region[] region;                  //


    //
	public String[] getDataValue(String[] wherestring)throws Exception{

		String ls_sql = "";

    ls_sql += " SELECT GETFIXLICS('"+wherestring[0].trim()+"','"+wherestring[1].trim()+"','"+wherestring[2].trim()+"') INFO FROM DUAL ";


       	JDBCConnection conn =  JDBCConnectionFactory.getJDBCConnection("SynctConn");
         System.err.println(ls_sql);

         //
         int li_total_row = 0;

         //
         Enumeration rows1 = null;
  
         int i = 0;
         rows1 = conn.getRows(ls_sql);
         conn.closeConnection();
   
         String [] rds=new String[20 ];
            String ldata = null;
         while( rows1 != null && rows1.hasMoreElements() ){
            DbRow row2 = (DbRow) rows1.nextElement();
           ldata  = Utils.convertToString(row2.get("INFO"));
            
         }
        rds = ldata.split(";");

         return rds;
	}

 	/**
  	*
  	*
  	*
  	*/
 	public synchronized  boolean outXLS(String userid,String[] wherestring) throws Exception{
        try{
        separator =  System.getProperty("file.separator");
	    	fs = new POIFSFileSystem(new FileInputStream(getPath() + "template" + separator + execlfilename));
	    	wb = new HSSFWorkbook(fs);
	    	sheet = wb.getSheetAt(0);
        sheet1 = wb.getSheetAt(1);
        sheet2 = wb.getSheetAt(2);
        sheet3 = wb.getSheetAt(3);        
	    	ps = sheet.getPrintSetup();
	    	sheet.setAutobreaks(false);
        sheet1.setAutobreaks(false);
        sheet2.setAutobreaks(false);
        sheet3.setAutobreaks(false);        
			execOut(userid,wherestring);
		}
		catch(Exception e ) {
		     throw new Exception(e);
		  	 //System.err.println("outXLS error is "+e);
			//return false;
		}
		return true;
	}

    //
    public void execOut(String userid,String[] wherestring) throws Exception{
      FileOutputStream fileOut = null;
      try {
            //
            System.err.println("bm10102.java: before getDataValue.");
            data=getDataValue(wherestring);
            System.err.println("bm10102.java: end getDataValue.");

             //
            body_style = copyPageBodyStyleBlock(sheet, 0,0,40,40);       // (int row, int start col, int cols)

            //
            body_value = copyPageBodyValueBlock(sheet, 0,0,40,40);

            //
            //region = copyMergedRegion(sheet);
            //System.err.println("1");
            //
        	//	pasteMergedRegion(sheet, region, (page_rows * page), 0);
    		pastePageBodyStyleBlock(sheet, body_style, 0, 0);
    		pastePageBodyValueBlock(sheet, body_value, 0, 0);
            printPageBody(0,data,0);

             //
            body_style = copyPageBodyStyleBlock(sheet2, 0,0,35,47);       // (int row, int start col, int cols)

            //
            body_value = copyPageBodyValueBlock(sheet2, 0,0,35,47);

            //
            pastePageBodyStyleBlock(sheet2, body_style, 0, 0);
            pastePageBodyValueBlock(sheet2, body_value, 0, 0);
                //
            printPageBody2(0,data,0); 
		
                //
	    		setPageBreak(ps);
	    	
            //
	        fileOut = new FileOutputStream(getPath() + "output" + separator + userid + execlfilename);
		    wb.write(fileOut);
      }catch(Exception e) {
         System.err.println("AP30000:execOut error is "+e.toString());
         throw new Exception(e.getMessage());
      }finally{
          fileOut.close();
      }
    }


    public void printPageBody(int j,String[] data1, int rowno) throws IOException {
     try{
         //
        HSSFRow row = sheet.getRow(2);
        HSSFCell cell1 =row.getCell((short)(0));
        setBig5CellValue(data1[0]+data1[1],cell1);  //A3     LICENSE_DESC

        row = sheet.getRow(4);
        cell1 =row.getCell((short)(0));   
        setBig5CellValue("申請人："+data1[2],cell1); //A5 APPL


        row = sheet.getRow(5);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("裝修地址："+data1[3],cell1);  //A6 ADDR

        row = sheet.getRow(6);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("裝修設計廠商："+data1[4],cell1);  //A7

        row = sheet.getRow(7);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("裝修施工廠商："+data1[5],cell1);  //A8

        row = sheet.getRow(8);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("審查機構："+data1[6],cell1);  //A9

        row = sheet.getRow(9);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("查驗人員："+data1[7],cell1);  //A10

        row = sheet.getRow(10);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("發證機關："+data1[8],cell1);  //A11


        }catch(Exception e)     {
            System.err.println("bm10102:printPageBody error is " + e);
         }

    }



    public void printPageBody2(int j,String[] data1, int rowno) throws IOException {
     try{
         //
        HSSFRow row = sheet2.getRow(2);
        HSSFCell cell1 =row.getCell((short)(0));
        setBig5CellValue(data1[0]+data1[1],cell1);  //A3     LICENSE_DESC

        row = sheet2.getRow(4);
        cell1 =row.getCell((short)(0));   
        setBig5CellValue("申請人："+data1[2],cell1); //A5 APPL


        row = sheet2.getRow(5);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("裝修地址："+data1[3],cell1);  //A6 ADDR

        row = sheet2.getRow(6);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("裝修設計廠商："+data1[4],cell1);  //A7

        row = sheet2.getRow(7);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("裝修施工廠商："+data1[5],cell1);  //A8

        row = sheet2.getRow(8);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("審查機構："+data1[6],cell1);  //A9

        row = sheet2.getRow(9);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("查驗人員："+data1[7],cell1);  //A10

        row = sheet2.getRow(10);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("發證機關："+data1[8],cell1);  //A11

        }catch(Exception e)     {
            System.err.println("bm10102:printPageBody error is " + e);
         }

    }





 }
