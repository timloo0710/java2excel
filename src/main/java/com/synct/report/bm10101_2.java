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

public class bm10101_2 extends Ole2Adapter {

	private		int onepage_detail = 22;     //一頁報表有幾列detail
	private    	int dtl_start_row = 0;      //detail從page裡的第幾列開始
	private    	int dtl_cols = 34;           //detail資料有幾欄
	private    	String execlfilename = "bm10101_2.xls";  //excel檔名

    public bm10101_2() {
        page_rows = 22;     //一頁報表有幾列
	}


    //畫面條件
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
	public HSSFSheet sheet,sheet1,sheet2,sheet3;
	public HSSFPrintSetup ps;

	HSSFCellStyle[][] header_style;   //header區塊的style
	HSSFCellStyle[][] body_style;     //body區塊的style

	String[][] data;
  String[] data_other;
	String[][] header_value;          //header區塊的欄位名稱,或標籤
	String[][] body_value;            //body區塊的欄位名稱,或標籤

	Region[] region;                  //合併儲存格陣列

  //依畫面條件從資料庫取得 雜項工作物 資料
  public String[] getData(String[] wherestring)throws Exception{
    String KEY = wherestring[0];
    String ls_sql = "";
    ls_sql += " select  Comb_Work(CONSNAME,BUILDING_KIND,LENGTH,HEIGHT,WIDE,AREA,CONNUM,DESE) other_work FROM BM_WORK ";  //
    ls_sql += " WHERE INDEX_KEY = '"+KEY+"'  order by PERSON_SEQ "; 

        JDBCConnection conn =  JDBCConnectionFactory.getJDBCConnection("SynctConn");
         System.out.println(ls_sql);

         String[] ds = new String[30];
         //抓取資料
         Enumeration rows1 = null;
         DbRow CurrentRecord;
         int i = 0;
         System.out.println("before getRows");         
         rows1 = conn.getRows(ls_sql);
         System.out.println("after getRows");         
         conn.closeConnection();
          while (rows1 != null && rows1.hasMoreElements()) {
            CurrentRecord = (DbRow) rows1.nextElement();
            ds[i] = Utils.convertToString(CurrentRecord.get("other_work"));
            System.out.println(Utils.convertToString(CurrentRecord.get("other_work")));
            i++;
          }


         return ds; 

  }  

    //依畫面條件從資料庫取得資料
	public String[][] getDataValue(String[] wherestring)throws Exception{

		String KEY = wherestring[0];
		//String KIND = wherestring[1];
    //String NO1 = wherestring[2];
    //String NO2 = wherestring[3];
    //String WORD = wherestring[4];

		String ls_sql = "";
    ls_sql += " SELECT B.BASE_AREA_ARC  AREA_ARC,B.BASE_AREA_SHRINK AREA_SHRINK,  B.BASE_AREA_OTHER AREA_OTHER,B.BASE_AREA_TOTAL AREA_TOTAL,B.PRICE PRICE,"; //5
    ls_sql += " B.LICENSE_DESC LICENSE_DESC,B.IDENTIFY_LICE_DATE IDENTIFY_LICE_DATE,";
    ls_sql += " B.COMMENCE_DATE COMMENCE_DATE,B.VALID_MONTH VALID_MONTH,B.APPROVE_LICE_DATE APPROVE_LICE_DATE,B.RECEIVE_LICE_DATE RECEIVE_LICE_DATE,"; //4
    ls_sql += " comb_use_category(B.USE_CATEGORY_CODE1,B.USE_CATEGORY_CODE2,B.USE_CATEGORY_CODE3) USE_CATEGORY_CODE_DESC,"; //1
    ls_sql += "  ( SELECT GETLANNO (DIST, SECTION, ROAD_NO1, ROAD_NO2) FROM BM_LAN WHERE INDEX_KEY= B.INDEX_KEY AND SPOKESMAN='Y' ) LANNO,"; //1
    ls_sql += "  (SELECT NAME FROM BM_P01 WHERE INDEX_KEY =B.INDEX_KEY AND SPOKESMAN='Y') P01_NAME,"; //1
    ls_sql += "  (SELECT COMB_ADDR1( O_ADDRADR_DESC,  O_ADDRAD1,  O_ADDRAD2,  O_ADDRAD3,  O_ADDRAD4,  O_ADDRAD5,  O_ADDRAD6,  O_ADDRAD6_1,  O_ADDRAD7,  O_ADDRAD7_1,  O_ADDRAD8)";
    ls_sql += "   FROM BM_P01 WHERE INDEX_KEY =B.INDEX_KEY AND SPOKESMAN='Y') ADDR,"; //1
    ls_sql += "   (SELECT NAME FROM BM_P02 WHERE INDEX_KEY=B.INDEX_KEY AND SPOKESMAN='Y') P02_NAME,"; //1
    ls_sql += "   (SELECT OFFICE_NAME FROM BM_P02 WHERE INDEX_KEY=B.INDEX_KEY AND SPOKESMAN='Y') OFFICE_NAME"; //1
    ls_sql += " from BM_BASE B ";
    ls_sql += " WHERE INDEX_KEY= '"+KEY+"' ";    

	
        //勾選
	//	if( !StringUtils.isEmpty(wherestring[1]) )
	//		ls_sql += " AND LMD_VOID = " + wherestring[1] ;
  //
	//	ls_sql += " ORDER BY SEQ ";



       	JDBCConnection conn =  JDBCConnectionFactory.getJDBCConnection("SynctConn");
         System.err.println(ls_sql);

         //抓取總筆數
         int li_total_row = 0;

         //抓取資料
         Enumeration rows1 = null;
         Enumeration rows2 = null;
         int i = 0;
         rows1 = conn.getRows(ls_sql);
         rows2 = conn.getRows(ls_sql);
         conn.closeConnection();
         //計算總筆數
         while( rows2 != null && rows2.hasMoreElements() ){
            DbRow row2 = (DbRow) rows2.nextElement();
            li_total_row++;
         }
         System.err.println("li_total_row="+li_total_row);
         String [][] rds=null;
         rds = new String[(int)li_total_row ][17];  //  row , column 
		 //double d_tot[] = new double[10];
         //給初始值
         for (int m=0;m<rds.length;m++){
         	for(int n=0;n<rds[m].length;n++){
         		rds[m][n] = "";
         	}
         }
         String AREA_ARC,AREA_SHRINK,AREA_OTHER,AREA_TOTAL,PRICE;
         String COMMENCE_DATE,VALID_MONTH,APPROVE_LICE_DATE,RECEIVE_LICE_DATE;
         String USE_CATEGORY_CODE_DESC,LANNO,P01_NAME,ADDR,P02_NAME,OFFICE_NAME;
         String LICENSE_DESC,IDENTIFY_LICE_DATE;
         while( rows1 != null && rows1.hasMoreElements() ){
            DbRow row2 = (DbRow) rows1.nextElement();
            
            AREA_ARC=Utils.convertToString(row2.get("AREA_ARC"));
            AREA_SHRINK=Utils.convertToString(row2.get("AREA_SHRINK"));
            AREA_OTHER=Utils.convertToString(row2.get("AREA_OTHER"));
            AREA_TOTAL=Utils.convertToString(row2.get("AREA_TOTAL"));
            PRICE=Utils.convertToString(row2.get("PRICE"));
            COMMENCE_DATE=Utils.convertToString(row2.get("COMMENCE_DATE"));
            VALID_MONTH=Utils.convertToString(row2.get("VALID_MONTH"));
            APPROVE_LICE_DATE=Utils.convertToString(row2.get("APPROVE_LICE_DATE"));
            RECEIVE_LICE_DATE=Utils.convertToString(row2.get("RECEIVE_LICE_DATE"));
            USE_CATEGORY_CODE_DESC=Utils.convertToString(row2.get("USE_CATEGORY_CODE_DESC"));
            LANNO=Utils.convertToString(row2.get("LANNO"));
            P01_NAME=Utils.convertToString(row2.get("P01_NAME"));
            ADDR=Utils.convertToString(row2.get("ADDR"));
            P02_NAME=Utils.convertToString(row2.get("P02_NAME"));
            OFFICE_NAME=Utils.convertToString(row2.get("OFFICE_NAME"));
            LICENSE_DESC=Utils.convertToString(row2.get("LICENSE_DESC"));
            IDENTIFY_LICE_DATE=Utils.convertToString(row2.get("IDENTIFY_LICE_DATE"));

            rds[i][0] = StringUtils.isEmpty(AREA_ARC) ? "* * *": AREA_ARC;
            rds[i][1]  = StringUtils.isEmpty(AREA_SHRINK) ? "* * *": AREA_SHRINK;
            rds[i][2]  = StringUtils.isEmpty(AREA_OTHER) ? "* * *": AREA_OTHER;
            rds[i][3]  = StringUtils.isEmpty(AREA_TOTAL) ? "* * *": AREA_TOTAL;


            if (StringUtils.isEmpty(PRICE))
              PRICE= "";
            else
              PRICE= String.format("%,d 元",Integer.parseInt(PRICE));
             
            rds[i][4]  =  PRICE;
            
            if (StringUtils.isEmpty(COMMENCE_DATE))
             COMMENCE_DATE= "";
            else
              COMMENCE_DATE= COMMENCE_DATE.substring(0,3)+"年"+COMMENCE_DATE.substring(3,5)+"月"+COMMENCE_DATE.substring(5,7)+"日";
            
            rds[i][5]  = COMMENCE_DATE;

            rds[i][6]  = StringUtils.isEmpty(VALID_MONTH) ? "": ("開工之日起   "+VALID_MONTH+ "  個月內完工");
            
            if (StringUtils.isEmpty(APPROVE_LICE_DATE))
             APPROVE_LICE_DATE= "";
            else
              APPROVE_LICE_DATE= APPROVE_LICE_DATE.substring(0,3)+"年"+APPROVE_LICE_DATE.substring(3,5)+"月"+APPROVE_LICE_DATE.substring(5,7)+"日";
            rds[i][7]  = APPROVE_LICE_DATE;

            if (StringUtils.isEmpty(RECEIVE_LICE_DATE))
             RECEIVE_LICE_DATE= "";
            else
              RECEIVE_LICE_DATE= RECEIVE_LICE_DATE.substring(0,3)+"年"+RECEIVE_LICE_DATE.substring(3,5)+"月"+RECEIVE_LICE_DATE.substring(5,7)+"日";

            rds[i][8]  = RECEIVE_LICE_DATE;

            rds[i][9]  = StringUtils.isEmpty(USE_CATEGORY_CODE_DESC) ? "": USE_CATEGORY_CODE_DESC;
            rds[i][10]  = StringUtils.isEmpty(LANNO) ? "": LANNO;
            rds[i][11]  = StringUtils.isEmpty(P01_NAME) ? "": P01_NAME;
            rds[i][12]  = StringUtils.isEmpty(ADDR) ? "": ADDR;
            rds[i][13]  = StringUtils.isEmpty(P02_NAME) ? "": P02_NAME;
            rds[i][14]  = StringUtils.isEmpty(OFFICE_NAME) ? "": OFFICE_NAME;
            rds[i][15]  = StringUtils.isEmpty(LICENSE_DESC) ? "": LICENSE_DESC;
            
            if (StringUtils.isEmpty(IDENTIFY_LICE_DATE))
             IDENTIFY_LICE_DATE= "";
            else
              IDENTIFY_LICE_DATE= IDENTIFY_LICE_DATE.substring(0,3)+"  年"+IDENTIFY_LICE_DATE.substring(3,5)+"  月"+IDENTIFY_LICE_DATE.substring(5,7)+"  日";

            rds[i][16]  = IDENTIFY_LICE_DATE;
         
            i ++;
         }
         return rds;
	}

    //填寫頁首
	private void printHeader(String[] wherestring) throws Exception{
		//寫出頁首
        HSSFRow pageRow = sheet.getRow(page * page_rows);
        HSSFCell pageCell = pageRow.getCell((short)1);

        pageRow = sheet.getRow(page * page_rows + 1);
        pageCell = pageRow.getCell((short)0);

        setBig5CellValue( "執照號碼：" + Utils.convertToString(DBTools.dLookUp("LM_LICNUM", "LICENSEMEMO", "SEQ="+wherestring[0], "SynctConn")),pageCell);
	}

    //填寫頁尾
	private void printFoot(String printpage) throws Exception{
		   //寫出頁尾
           HSSFRow pageRow = sheet.getRow((page * page_rows) + 13);
           HSSFCell pageCell = pageRow.getCell((short)0);
		   setBig5CellValue(printpage ,pageCell);

	}


 	/**
  	*<br>目的：輸出報表
  	*<br>參數： 無
  	*<br>傳回：boolean
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

    //產生Excel檔
    public void execOut(String userid,String[] wherestring) throws Exception{
      FileOutputStream fileOut = null;
      try {
            //進資料庫查詢
            System.err.println("bm10101_2.java: before getDataValue.");
            data=getDataValue(wherestring);
            data_other=getData(wherestring);
            System.err.println("bm10101_2.java: end getDataValue.");

            //複製表頭樣式
 	        //header_style = copyPageHeaderStyle(sheet, 0,0,dtl_start_row,dtl_cols); // (int start_row, int start_col, int rows 到detail開始為止為header的列數, int cols)

 	        //複製表頭儲存格值
            //header_value = copyPageHeaderValue(sheet, 0,0,dtl_start_row,dtl_cols);

            //複製表身樣式
            body_style = copyPageBodyStyleBlock(sheet, 0,0,34,dtl_cols);       // (int row, int start col, int cols)

            //複製表身儲存格值
            body_value = copyPageBodyValueBlock(sheet, 0,0,34,dtl_cols);

            //複製第一頁報表內的合併儲存格
            //region = copyMergedRegion(sheet);

            //    pasteMergedRegion(sheet, region, 0, 0);
                pastePageBodyStyleBlock(sheet, body_style, 0, 0);
                pastePageBodyValueBlock(sheet, body_value, 0, 0);
                //填格子
                printPageBody(0,data[0],0); 


            //複製第二頁
            body_style = copyPageBodyStyleBlock(sheet1, 0,0,35,54);       // (int row, int start col, int cols)

            //複製表身儲存格值
            body_value = copyPageBodyValueBlock(sheet1, 0,0,35,54);

            //複製第二頁報表內的合併儲存格
            //region = copyMergedRegion(sheet1);

            //    pasteMergedRegion(sheet1, region, 0, 0);
                pastePageBodyStyleBlock(sheet1, body_style, 0, 0);
                pastePageBodyValueBlock(sheet1, body_value, 0, 0);
                //填格子
                printPageBody1(0,data[0],0); 



            //複製第三頁
            body_style = copyPageBodyStyleBlock(sheet2, 0,0,35,47);       // (int row, int start col, int cols)

            //複製表身儲存格值
            body_value = copyPageBodyValueBlock(sheet2, 0,0,35,47);

            //複製第三頁報表內的合併儲存格
            //region = copyMergedRegion(sheet2);
            //    pasteMergedRegion(sheet, region, 0, 0);
                pastePageBodyStyleBlock(sheet2, body_style, 0, 0);
                pastePageBodyValueBlock(sheet2, body_value, 0, 0);
                //填格子
                printPageBody2(0,data[0],0); 



            //複製第四頁
            body_style = copyPageBodyStyleBlock(sheet3, 0,0,35,54);       // (int row, int start col, int cols)

            //複製表身儲存格值
            body_value = copyPageBodyValueBlock(sheet3, 0,0,35,54);

            //複製第四頁報表內的合併儲存格
            //region = copyMergedRegion(sheet3);

            //    pasteMergedRegion(sheet3, region, 0, 0);
                pastePageBodyStyleBlock(sheet3, body_style, 0, 0);
                pastePageBodyValueBlock(sheet3, body_value, 0, 0);
                //填格子
                printPageBody3(0,data[0],0); 



            //先計算出總頁數,迴圈中針對每頁處理塞值的動作,也可以用總筆數去跑迴圈
            int total_page = 0;
            total_page=((data.length - 1)/onepage_detail) + 1;

            //每頁頁尾的總計值
            int total=0;
            int totalCount=0;

 


            //轉出Excel檔
	        fileOut = new FileOutputStream(getPath() + "output" + separator + userid + execlfilename);
		    wb.write(fileOut);
      }catch(Exception e) {
         System.err.println("AP30000:execOut error is "+e.toString());
         throw new Exception(e.getMessage());
      }finally{
          fileOut.close();
      }
    }


    public void printPageBody(int k,String[] data1, int rowno) throws IOException {
       try{
 		    //加入資料
        HSSFRow row = sheet.getRow(0);
        HSSFCell cell1 =row.getCell((short)(20));
        setBig5CellValue(data1[15],cell1);  //U1 LICENSE_DESC

 		    row = sheet.getRow(1);
        cell1 =row.getCell((short)(6));   
        setBig5CellValue(data1[11],cell1); //G2 P01_NAME


     	  row = sheet.getRow(2);
        cell1 =row.getCell((short)(6));
       	setBig5CellValue(data1[12],cell1);  //G3 ADDR

        row = sheet.getRow(3);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[13],cell1);  //G4 P02_NAME

        row = sheet.getRow(3);
        cell1 =row.getCell((short)(23));
        setBig5CellValue(data1[14],cell1);  //X4 OFFICE_NAME

        row = sheet.getRow(4);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[10],cell1);  //G5 LANNO

        row = sheet.getRow(5);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[12],cell1);  //G6 ADDR

        row = sheet.getRow(6);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[9],cell1);  //G7 USE_CATEGORY_CODE_DESC

        row = sheet.getRow(7);
        cell1 =row.getCell((short)(9));
        setBig5CellValue(data1[0]+"㎡",cell1);  //J8 AREA_ARC

        row = sheet.getRow(7);
        cell1 =row.getCell((short)(23));
        setBig5CellValue(data1[2]+"㎡",cell1);  //X8 AREA_OTHER

        row = sheet.getRow(8);
        cell1 =row.getCell((short)(9));
        setBig5CellValue(data1[1]+"㎡",cell1);  //J9 AREA_SHRINK

        row = sheet.getRow(8);
        cell1 =row.getCell((short)(23));
        setBig5CellValue(data1[3]+"㎡",cell1);  //X9 AREA_TOTAL

        row = sheet.getRow(9);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[4],cell1);  //G10 PRICE


        row = sheet.getRow(10);
        cell1 =row.getCell((short)(24));
        setBig5CellValue(data1[6],cell1);  //Y11 VALID_MONTH

        row = sheet.getRow(11);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[7],cell1);  //G12 APPROVE_LICE_DATE

        row = sheet.getRow(11);
        cell1 =row.getCell((short)(24));
        setBig5CellValue(data1[8],cell1);  //Y12 RECEIVE_LICE_DATE

        row = sheet.getRow(36);
        cell1 =row.getCell((short)(1));
        setBig5CellValue("上給    "+data1[11],cell1);  //B37 P01_NAME
        

        row = sheet.getRow(40);
        cell1 =row.getCell((short)(15));
        setBig5CellValue(data1[16],cell1);  //P41 IDENTIFY_LICE_DATE
 
        row = sheet.getRow(12);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("雜項工作物：",cell1); //A13
        int j=0;
        for(int i=0;i<data_other.length;i++) {  //A14 A15 A16 A17 A18
          row = sheet.getRow(13+i);
          cell1 =row.getCell((short)(0));
          setBig5CellValue(data_other[i],cell1); 
          j=i;
          if (StringUtils.isEmpty(data_other[i]))
            break;

          }

        row = sheet.getRow(15+j);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("以下空白",cell1); //

       }catch(Exception e)     {
          System.err.println("bm10101_2:printPageBody error is " + e);
       }

    }

    public void printPageBody1(int k,String[] data1, int rowno) throws IOException {
      try{

        HSSFRow row = sheet1.getRow(1);
        HSSFCell  cell1 =row.getCell((short)(0));
        setBig5CellValue("雜項工作物：",cell1); //A2

        row = sheet1.getRow(0);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[15],cell1);               //V1

        int j=0;
        for(int i=0;i<data_other.length;i++) {  //A14 A15 A16 A17 A18
          row = sheet1.getRow(2+i);
          cell1 =row.getCell((short)(0));
          setBig5CellValue(data_other[i],cell1); 
          j=i;
          if (StringUtils.isEmpty(data_other[i]))
            break;

          }

        row = sheet1.getRow(3+j);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("以下空白",cell1); //        

      }catch(Exception e)     {
          System.err.println("bm10101_2:printPageBody error is " + e);
      }

    }

public void printPageBody2(int k,String[] data1, int rowno) throws IOException {
       try{
        //加入資料
        HSSFRow row = sheet2.getRow(1);
        HSSFCell cell1 =row.getCell((short)(6));   
        setBig5CellValue(data1[11],cell1); //G2 P01_NAME

        row = sheet2.getRow(0);
        cell1 =row.getCell((short)(20));
        setBig5CellValue(data1[15],cell1);  //U1 LICENSE_DESC

        row = sheet2.getRow(2);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[12],cell1);  //G3 ADDR

        row = sheet2.getRow(3);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[13],cell1);  //G4 P02_NAME

        row = sheet2.getRow(3);
        cell1 =row.getCell((short)(23));
        setBig5CellValue(data1[14],cell1);  //X4 OFFICE_NAME

        row = sheet2.getRow(4);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[10],cell1);  //G5 LANNO

        row = sheet2.getRow(5);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[12],cell1);  //G6 ADDR

        row = sheet2.getRow(6);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[9],cell1);  //G7 USE_CATEGORY_CODE_DESC

        row = sheet2.getRow(7);
        cell1 =row.getCell((short)(9));
        setBig5CellValue(data1[0]+"㎡",cell1);  //J8 AREA_ARC

        row = sheet2.getRow(7);
        cell1 =row.getCell((short)(23));
        setBig5CellValue(data1[2]+"㎡",cell1);  //X8 AREA_OTHER

        row = sheet2.getRow(8);
        cell1 =row.getCell((short)(9));
        setBig5CellValue(data1[1]+"㎡",cell1);  //J9 AREA_SHRINK

        row = sheet2.getRow(8);
        cell1 =row.getCell((short)(23));
        setBig5CellValue(data1[3]+"㎡",cell1);  //X9 AREA_TOTAL

        row = sheet2.getRow(9);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[4],cell1);  //G10 PRICE


        row = sheet2.getRow(10);
        cell1 =row.getCell((short)(24));
        setBig5CellValue(data1[6],cell1);  //Y11 VALID_MONTH

        row = sheet2.getRow(11);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[7],cell1);  //G12 APPROVE_LICE_DATE

        row = sheet2.getRow(11);
        cell1 =row.getCell((short)(24));
        setBig5CellValue(data1[8],cell1);  //Y12 RECEIVE_LICE_DATE

        //row = sheet2.getRow(36);
        //cell1 =row.getCell((short)(1));
        //setBig5CellValue("上給    "+data1[11],cell1);  //B37 P01_NAME
        

        //row = sheet2.getRow(40);
        //cell1 =row.getCell((short)(15));
        //setBig5CellValue(data1[16],cell1);  //P41 IDENTIFY_LICE_DATE
 
        row = sheet2.getRow(12);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("雜項工作物：",cell1); //A13
        int j=0;
        for(int i=0;i<data_other.length;i++) {  //A14 A15 A16 A17 A18
          row = sheet2.getRow(13+i);
          cell1 =row.getCell((short)(0));
          setBig5CellValue(data_other[i],cell1); 
          j=i;
          if (StringUtils.isEmpty(data_other[i]))
            break;
          }

        row = sheet2.getRow(15+j);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("以下空白",cell1); //

       }catch(Exception e)     {
          System.err.println("sheet2:printPageBody error is " + e);
       }

    }

    public void printPageBody3(int k,String[] data1, int rowno) throws IOException {
      try{

        HSSFRow row = sheet3.getRow(1);
        HSSFCell  cell1 =row.getCell((short)(0));
        setBig5CellValue("雜項工作物：",cell1); //A2

        row = sheet3.getRow(0);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[15],cell1);               //V1

        int j=0;
        for(int i=0;i<data_other.length;i++) {  //A14 A15 A16 A17 A18
          row = sheet3.getRow(2+i);
          cell1 =row.getCell((short)(0));
          setBig5CellValue(data_other[i],cell1); 
          j=i;
           if (StringUtils.isEmpty(data_other[i]))
            break;

          }

        row = sheet3.getRow(3+j);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("以下空白",cell1); //        

      }catch(Exception e)     {
          System.err.println("sheet3:printPageBody error is " + e);
      }

    }


 }

        //加入資料

/*
0 1 2 3 4 5 6 7 8 910111213141516171819202122232425 6 7 8 930 1 2 3 4 5 
A B C D E F G H I J K L M N O P Q R S T U V W X Y ZAAABACADAEAFAGAHAIAJ
        rds[i][0] = StringUtils.isEmpty(AREA_ARC) ? "null": AREA_ARC;
        rds[i][1]  = StringUtils.isEmpty(AREA_SHRINK) ? "null": AREA_SHRINK;
        rds[i][2]  = StringUtils.isEmpty(AREA_OTHER) ? "null": AREA_OTHER;
        rds[i][3]  = StringUtils.isEmpty(AREA_TOTAL) ? "null": AREA_TOTAL;
        rds[i][4]  = StringUtils.isEmpty(PRICE) ? "null": PRICE;
        rds[i][5]  = StringUtils.isEmpty(COMMENCE_DATE) ? "null": COMMENCE_DATE;
        rds[i][6]  = StringUtils.isEmpty(VALID_MONTH) ? "null": VALID_MONTH;
        rds[i][7]  = StringUtils.isEmpty(APPROVE_LICE_DATE) ? "null": APPROVE_LICE_DATE;
        rds[i][8]  = StringUtils.isEmpty(RECEIVE_LICE_DATE) ? "null": RECEIVE_LICE_DATE;
        rds[i][9]  = StringUtils.isEmpty(USE_CATEGORY_CODE_DESC) ? "null": USE_CATEGORY_CODE_DESC;
        rds[i][10]  = StringUtils.isEmpty(LANNO) ? "null": LANNO;
        rds[i][11]  = StringUtils.isEmpty(P01_NAME) ? "null": P01_NAME;
        rds[i][12]  = StringUtils.isEmpty(ADDR) ? "null": ADDR;
        rds[i][13]  = StringUtils.isEmpty(P02_NAME) ? "null": P02_NAME;
        rds[i][14]  = StringUtils.isEmpty(OFFICE_NAME) ? "null": OFFICE_NAME;

        rds[i][15]LICENSE_DESC=Utils.convertToString(row2.get("LICENSE_DESC"));
        rds[i][16]IDENTIFY_LICE_DATE=Utils.convertToString(row2.get("IDENTIFY_LICE_DATE"));


 */


