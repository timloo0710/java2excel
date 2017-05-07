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

public class bm10104_2 extends Ole2Adapter {

	private		int onepage_detail = 20000;     //一頁報表有幾列detail
	private    	int dtl_start_row = 1;      //detail從page裡的第幾列開始
	private    	int dtl_cols = 39;           //detail資料有幾欄
	private    	String execlfilename = "BM10104_2.xls";  //excel檔名

    public bm10104_2() {
        page_rows = 20000;     //一頁報表有幾列
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
	public HSSFSheet sheet;
	public HSSFPrintSetup ps;

	HSSFCellStyle[][] header_style;   //header區塊的style
	HSSFCellStyle[][] body_style;     //body區塊的style

	String[][] data;
	String[][] header_value;          //header區塊的欄位名稱,或標籤
	String[][] body_value;            //body區塊的欄位名稱,或標籤

	Region[] region;                  //合併儲存格陣列


    //依畫面條件從資料庫取得資料
	public String[][] getDataValue(String[] wherestring)throws Exception{
		String ls_sql = "";

		ls_sql += " SELECT SEQ, LM_YY, LM_WORD, LM_TYPE, LM_NO1, LM_NO2 ";
		ls_sql += " FROM  LICENSEMEMO";
		ls_sql += " WHERE  1 = 1 " ;

		if( !StringUtils.isEmpty(wherestring[0]) )
			ls_sql += " AND (RR_REGNUM LIKE '%" + wherestring[0] + "%' OR SEQ IN (SELECT LM_SEQ FROM LICENSEMEMO_DE WHERE LMD_MEMO LIKE '%" + wherestring[0] + "%'))";

		if( !StringUtils.isEmpty(wherestring[1]) )
			ls_sql += " AND LM_YY = '" + wherestring[1]  + "'";

		if( !StringUtils.isEmpty(wherestring[2]) )
			ls_sql += " AND LM_NO1 = '" + wherestring[2]  + "'";

		if( !StringUtils.isEmpty(wherestring[3]) )
			ls_sql += " AND LM_NO2 = '" + wherestring[3]  + "'";

		if( !StringUtils.isEmpty(wherestring[4]) )
			ls_sql += " AND LM_WORD = '" + wherestring[4]  + "'";

		if( !StringUtils.isEmpty(wherestring[5]) )
			ls_sql += " AND LM_TYPE = '" + wherestring[5] + "'";

		ls_sql += " ORDER BY 2,3,4,5 ";
//System.err.println("***************ls_sql=" + ls_sql);


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
         String INDEX_KEY ="";
         String REG_YY    ="";
         String REG_NO    ="";
         String REG_CG    ="";
         rds = new String[(int)li_total_row ][39];
		 //double d_tot[] = new double[10];
         //給初始值
         for (int m=0;m<rds.length;m++){
         	for(int n=0;n<rds[m].length;n++){
         		rds[m][n] = "";
         	}
         }

         while( rows1 != null && rows1.hasMoreElements() ){
            DbRow row2 = (DbRow) rows1.nextElement();
            
             String LM_KIND = "";
             
             if ( !StringUtils.isEmpty(Utils.convertToString(row2.get("LM_TYPE"))) &&  Utils.convertToString(row2.get("LM_TYPE")).equals("01"))
                LM_KIND = "1";
             else if  ( !StringUtils.isEmpty(Utils.convertToString(row2.get("LM_TYPE"))) &&  Utils.convertToString(row2.get("LM_TYPE")).equals("05"))
                LM_KIND = "2";  
             else if  ( !StringUtils.isEmpty(Utils.convertToString(row2.get("LM_TYPE"))) &&  Utils.convertToString(row2.get("LM_TYPE")).equals("11"))
                LM_KIND = "3";  
             else if  ( !StringUtils.isEmpty(Utils.convertToString(row2.get("LM_TYPE"))) &&  Utils.convertToString(row2.get("LM_TYPE")).equals("03"))
                LM_KIND = "4";  
             
             
               
             rds[i][0] = Utils.convertToString(i+1);
             rds[i][2]  = Utils.convertToString(row2.get("LM_YY")) + Utils.convertToString(DBTools.dLookUp("CODENAME", "PARA", "KIND = 'LM_WORD' AND CODE ='"+Utils.convertToString(row2.get("LM_WORD")) + "'", "SynctConn")) + Utils.convertToString(DBTools.dLookUp("CODENAME", "PARA", "KIND = 'LM_TYPE' AND CODE ='"+ Utils.convertToString(row2.get("LM_TYPE")) + "'", "SynctConn")) + "第" + Utils.convertToString(row2.get("LM_NO1")) + "號";
            
		     ls_sql = "SELECT LMD_MEMO FROM LICENSEMEMO_DE WHERE LMD_VOID = '1' AND LM_SEQ = " + Utils.convertToString(row2.get("SEQ"));
			 if( !StringUtils.isEmpty(wherestring[0]) )
				 ls_sql += " AND LMD_MEMO LIKE '%" + wherestring[0] + "%'";

             String s_LMD_MEMO = "";
             
	       	 conn =  JDBCConnectionFactory.getJDBCConnection("SynctConn");
	         Enumeration rows3 = conn.getRows(ls_sql);
	         while( rows3 != null && rows3.hasMoreElements() ){
                 DbRow row3 = (DbRow) rows3.nextElement();
				 if( !StringUtils.isEmpty(s_LMD_MEMO))
                    s_LMD_MEMO += "\n" +  Utils.convertToString(row3.get("LMD_MEMO"));
                 else
                    s_LMD_MEMO += Utils.convertToString(row3.get("LMD_MEMO"));
             }          
            
             rds[i][1]  = s_LMD_MEMO;

		     ls_sql = "SELECT INDEX_KEY,LICENSE_DESC,USE_CATEGORY_CODE_DESC,IDENTIFY_LICE_DATE,RECEIVE_LICE_DATE ";
		     ls_sql += " ,UP_FLOOR_NO,DN_FLOOR_NO,TOT_HOUSE_NO,BUILDING_HEIGHT,TOTAL_CONSTRU_AREA,PRICE ";
		     ls_sql += " ,BASE_AREA_OTHER,STATUTORY_OPEN_SPACE,LAW_COVER_RATE,LAW_SPACE_RATE,USAGE_CODE_DESC ";
		     ls_sql += " ,PARK_SUM1,PARK_SUM2,PARK_SUM3,COMMENCE_DATE,COMPLETE_DATE,PUBLIC_CODE,BASE_AREA_PURPOSE ";
		     ls_sql += " ,LICENSE_YY_OLD,LICENSE_KIND_OLD,LICENSE_NO1_OLD,LICENSE_NO2_OLD,LICENSE_WORD_OLD ";
		     ls_sql += " ,LICENSE_DESC_OLD,REG_YY,REG_NO,REG_CG ";
		     ls_sql += " FROM BM_BASE ";
		     ls_sql += " WHERE LICENSE_YY = '" + Utils.convertToString(row2.get("LM_YY")) + "'";
		     ls_sql += " AND LICENSE_KIND = '" + LM_KIND + "'";
		     ls_sql += " AND LICENSE_NO1  = '" + Utils.convertToString(row2.get("LM_NO1")) + "'";
		     ls_sql += " AND LICENSE_NO2  = '" + Utils.convertToString(row2.get("LM_NO2")) + "'";
		     ls_sql += " AND LICENSE_WORD = '" + Utils.convertToString(row2.get("LM_WORD")) + "'";

//System.err.println("***************2");
//System.err.println("***************ls_sql=" + ls_sql);

	         rows3 = conn.getRows(ls_sql);
	         while( rows3 != null && rows3.hasMoreElements() ){
                 DbRow row3 = (DbRow) rows3.nextElement();
                 INDEX_KEY = Utils.convertToString(row3.get("INDEX_KEY"));
                 rds[i][2] = Utils.convertToString(row3.get("LICENSE_DESC"));
                 rds[i][6] = Utils.convertToString(row3.get("USE_CATEGORY_CODE_DESC"));
                 rds[i][7] = Utils.convertToString(row3.get("IDENTIFY_LICE_DATE"));
                 rds[i][8] = Utils.convertToString(row3.get("RECEIVE_LICE_DATE"));
                 rds[i][9] = Utils.convertToString(row3.get("UP_FLOOR_NO"));
                 rds[i][10] = Utils.convertToString(row3.get("DN_FLOOR_NO"));
                 rds[i][11] = Utils.convertToString(row3.get("TOT_HOUSE_NO"));
                 rds[i][12] = Utils.convertToString(row3.get("BUILDING_HEIGHT"));
                 rds[i][13] = Utils.convertToString(row3.get("TOTAL_CONSTRU_AREA"));
                 rds[i][14] = Utils.convertToString(row3.get("PRICE"));
                 rds[i][15] = Utils.convertToString(row3.get("BASE_AREA_OTHER"));
                 rds[i][16] = Utils.convertToString(row3.get("STATUTORY_OPEN_SPACE"));
                 rds[i][17] = Utils.convertToString(row3.get("LAW_COVER_RATE"));
                 rds[i][18] = Utils.convertToString(row3.get("LAW_SPACE_RATE"));
                 rds[i][19] = Utils.convertToString(row3.get("USAGE_CODE_DESC"));
                 rds[i][24] = Utils.convertToString(row3.get("PARK_SUM1"));
                 rds[i][25] = Utils.convertToString(row3.get("PARK_SUM2"));
                 rds[i][26] = Utils.convertToString(row3.get("PARK_SUM3"));
                 rds[i][27] = Utils.convertToString(row3.get("COMMENCE_DATE"));
                 rds[i][28] = Utils.convertToString(row3.get("COMPLETE_DATE"));
                 rds[i][29] = Utils.convertToString(row3.get("PUBLIC_CODE"));
                 rds[i][30] = Utils.convertToString(row3.get("BASE_AREA_PURPOSE"));
                 rds[i][31] = Utils.convertToString(row3.get("LICENSE_DESC_OLD"));
                 
//System.err.println("***************3");
                 
                 String ls_where  = " LICENSE_YY = '" + Utils.convertToString(row3.get("LICENSE_YY_OLD")) + "'";
		         ls_where += " AND LICENSE_KIND = '" + Utils.convertToString(row3.get("LICENSE_KIND_OLD")) + "'";
		         ls_where += " AND LICENSE_NO1  = '" + Utils.convertToString(row3.get("LICENSE_NO1_OLD")) + "'";
		         ls_where += " AND LICENSE_NO2  = '" + Utils.convertToString(row3.get("LICENSE_NO2_OLD")) + "'";
		         ls_where += " AND LICENSE_WORD = '" + Utils.convertToString(row3.get("LICENSE_WORD_OLD")) + "'";

                 rds[i][32] = Utils.convertToString(DBTools.dLookUp("COMMENCE_DATE", "BM_BASE", ls_where, "SynctConn"));
                 
                 REG_YY = Utils.convertToString(row3.get("REG_YY"));
                 REG_NO = Utils.convertToString(row3.get("REG_NO"));
                 REG_CG = Utils.convertToString(row3.get("REG_CG"));

//System.err.println("***************4");
                 if (!StringUtils.isEmpty(REG_YY))
                    rds[i][33] = REG_YY + "-" + REG_NO + "-" + REG_CG;
                    
             }          

		     ls_sql = "SELECT SPOKESMAN, PERSON_SEQ, NAME, Comb_Addr1(addradr_desc,addrad1,addrad2,addrad3,addrad4,addrad5,addrad6,addrad6_1,addrad7,addrad7_1,addrad8) ADDR ";
             ls_sql += "  FROM BM_P01 ";
             ls_sql += "  WHERE ROWNUM=1 AND INDEX_KEY = '" + INDEX_KEY +"' ORDER BY 1,2";
//System.err.println("***************5");
//System.err.println("***************ls_sql=" + ls_sql);

	         rows3 = conn.getRows(ls_sql);
	         while( rows3 != null && rows3.hasMoreElements() ){
                 DbRow row3 = (DbRow) rows3.nextElement();
                 rds[i][3]  = Utils.convertToString(row3.get("NAME"));
                 rds[i][4]  = Utils.convertToString(row3.get("ADDR"));
                 rds[i][20] = Utils.convertToString(row3.get("ADDR"));
             }          

//System.err.println("***************6");

		     ls_sql = "SELECT SPOKESMAN, PERSON_SEQ, NAME ";
             ls_sql += "  FROM BM_P02 ";
             ls_sql += "  WHERE ROWNUM=1 AND INDEX_KEY = '" + INDEX_KEY +"' ORDER BY 1,2";

//System.err.println("***************7");
//System.err.println("***************ls_sql=" + ls_sql);
	         rows3 = conn.getRows(ls_sql);
	         while( rows3 != null && rows3.hasMoreElements() ){
                 DbRow row3 = (DbRow) rows3.nextElement();
                 rds[i][21] = Utils.convertToString(row3.get("NAME"));
             }          
//System.err.println("***************8");
//System.err.println("***************ls_sql=" + ls_sql);

		     ls_sql = "SELECT SPOKESMAN, PERSON_SEQ, NAME ";
             ls_sql += "  FROM BM_P03 ";
             ls_sql += "  WHERE ROWNUM=1 AND INDEX_KEY = '" + INDEX_KEY +"' ORDER BY 1,2";

	         rows3 = conn.getRows(ls_sql);
	         while( rows3 != null && rows3.hasMoreElements() ){
                 DbRow row3 = (DbRow) rows3.nextElement();
                 rds[i][22] = Utils.convertToString(row3.get("NAME"));
             }          
//System.err.println("***************9");

		     ls_sql = "SELECT SPOKESMAN, PERSON_SEQ, COMPANY_NAME || BOSS BOSS ";
             ls_sql += "  FROM BM_P04 ";
             ls_sql += "  WHERE ROWNUM=1 AND INDEX_KEY = '" + INDEX_KEY +"' ORDER BY 1,2";

//System.err.println("***************10");
//System.err.println("***************ls_sql=" + ls_sql);

	         rows3 = conn.getRows(ls_sql);
	         while( rows3 != null && rows3.hasMoreElements() ){
                 DbRow row3 = (DbRow) rows3.nextElement();
                 rds[i][23] = Utils.convertToString(row3.get("BOSS"));
             }          
//System.err.println("***************11");

		     ls_sql = "SELECT REG_KIND,IO_DATE,WORK_DAYS,OT_DAYS,DELAY_DAYS ";
		     ls_sql += "  ,CASE WHEN END_DATE3 IS NOT NULL THEN END_DATE3 ELSE CASE WHEN END_DATE2 IS NOT NULL THEN END_DATE2 ELSE END_DATE2 END END END_DATE ";
             ls_sql += "  FROM BMSREGT ";
             ls_sql += "  WHERE ROWNUM=1 AND REG_YY = '" + REG_YY +"' ";
             ls_sql += "  AND REG_NO = '" + REG_NO +"' ";
             ls_sql += "  AND REG_CG = '" + REG_CG +"' ";
//System.err.println("***************12");
//System.err.println("***************ls_sql=" + ls_sql);

	         rows3 = conn.getRows(ls_sql);
	         while( rows3 != null && rows3.hasMoreElements() ){
                 DbRow row3 = (DbRow) rows3.nextElement();
                 rds[i][5]  = Utils.convertToString(DBTools.dLookUp("CODE_DESC", "BLDCODE", "CODE_TYPE = 'OFC' AND CODE_SEQ='"+ Utils.convertToString(row3.get("REG_KIND")) + "'", "SynctConn")) ;
                 rds[i][34] = Utils.convertToString(row3.get("IO_DATE"));
                 rds[i][35] = Utils.convertToString(row3.get("WORK_DAYS"));
                 rds[i][36] = Utils.convertToString(row3.get("OT_DAYS"));
                 rds[i][37] = Utils.convertToString(row3.get("DELAY_DAYS"));
                 rds[i][38] = Utils.convertToString(row3.get("END_DATE"));
             }          
//System.err.println("***************13");
	         conn.closeConnection();




//System.err.println("rds[i][1]=" + rds[i][1]);
            
            
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
	    	ps = sheet.getPrintSetup();
	    	sheet.setAutobreaks(false);
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
            System.err.println("bm10104_2.java: before getDataValue.");
            data=getDataValue(wherestring);
            System.err.println("bm10104_2.java: end getDataValue.");

            //複製表頭樣式
 	        //header_style = copyPageHeaderStyle(sheet, 0,0,dtl_start_row,dtl_cols); // (int start_row, int start_col, int rows 到detail開始為止為header的列數, int cols)

 	        //複製表頭儲存格值
            //header_value = copyPageHeaderValue(sheet, 0,0,dtl_start_row,dtl_cols);

            //複製表身樣式
            body_style = copyPageBodyStyleBlock(sheet, 1,0,2,dtl_cols);       // (int row, int start col, int cols)

            //複製表身儲存格值
            body_value = copyPageBodyValueBlock(sheet, 1,0,2,dtl_cols);

            //複製第一頁報表內的合併儲存格
            region = copyMergedRegion(sheet);
            //System.err.println("1");
            //先計算出總頁數,迴圈中針對每頁處理塞值的動作,也可以用總筆數去跑迴圈
            int total_page = 0;
            total_page=((data.length - 1)/onepage_detail) + 1;

            //每頁頁尾的總計值
            int total=0;
            int totalCount=0;

    		for(int i=0;i<total_page;i++) {
    			//先貼上之前複製的資料  ,在此新頁中加入header
    			if(page != 0) {
					//pastePageHeaderStyle(sheet, header_style, (page_rows * page), 0);
					//pastePageHeaderValue(sheet, header_value, (page_rows * page), 0);
		    		pasteMergedRegion(sheet, region, (page_rows * page), 0);
		    		pastePageBodyStyleBlock(sheet, body_style, (page_rows * page  + dtl_start_row), 0);
		    		pastePageBodyValueBlock(sheet, body_value, (page_rows * page) + dtl_start_row, 0);
		    	}
    			//填頁首的值
//System.err.println("bi30101.java: before printHeader.");
                //printHeader(wherestring);
//System.err.println("bi30101.java: end printHeader.");
	    		//寫入detail
	      		for(int j=0;j<onepage_detail;j++) {
//System.err.println("bi30101.java: page =" + page);
                        if(data.length>onepage_detail * page + j){

//System.err.println("bi30101.java: pastePageBodyStyleBlock start" );
				    		//pasteMergedRegion(sheet, region, j , 0);
				    		pastePageBodyStyleBlock(sheet, body_style, j  + dtl_start_row, 0);
				    		pastePageBodyValueBlock(sheet, body_value, j  + dtl_start_row, 0);
				    		
				    		
//System.err.println("bi30101.java: pastePageBodyStyleBlock end" );
         	    			printPageBody(onepage_detail * page + j + 1,data[onepage_detail * page + j],(page_rows * page) + dtl_start_row + j * 1);
     	    			}else{
            			    break;
     	    			}

	    		}
	    		//填頁尾的值
	    		//printFoot("第" + (page + 1) + "頁，共" + (total_page) + "頁");

                //換頁
	    		setPageBreak(ps);
	    	}
            //轉出Excel檔
	        fileOut = new FileOutputStream(getPath() + "output" + separator + userid + execlfilename);
		    wb.write(fileOut);
      }catch(Exception e) {
         System.err.println("AP30002:execOut error is "+e.toString());
         throw new Exception(e.getMessage());
      }finally{
          fileOut.close();
      }
    }


    public void printPageBody(int j,String[] data1, int rowno) throws IOException {
               try{
         		    //加入資料
         		    HSSFRow row = sheet.getRow(rowno);
                    HSSFCell cell1 =null;
                    
                    for (int g = 0 ;g <=38 ; g++){
	             	    cell1 = row.getCell((short)(g));
	               	    setBig5CellValue(data1[g],cell1);
                    }


               }catch(Exception e)     {
                  System.err.println("bi30101:printPageBody error is " + e);
               }

    }


 }