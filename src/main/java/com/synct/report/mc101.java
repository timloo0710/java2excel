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

public class mc101 extends Ole2Adapter {

	private		int onepage_detail = 20000;     //一頁報表有幾列detail
	private    	int dtl_start_row = 1;          //detail從page裡的第幾列開始
	private    	int dtl_cols = 19;              //detail資料有幾欄
	private    	String execlfilename = "MC101.xls";  //excel檔名

    public mc101() {
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

		ls_sql += " SELECT REG_YY || '-' || REG_NO || '-' || REG_CG REG_NO, REG_KIND, IO_DATE, IO_DEP3_NAME, LAST_RESULT_C, TMDNUM, LICENSE_YY || '-' || LICENSE_NO1 LICENSE_NO1, WORK_DAYS, ACTN_EMPL, case when RELATION1 = '5' then NAME1 else case when RELATION2 = '5' then NAME2 else case when RELATION3 = '5' then NAME3 end end end AS RELATION, REAL_WORK_DAYS, UP_FLOOR_NO, DN_FLOOR_NO, PUBLIC_FLAG, CLOSE_DATE, AREA_FLAG, P01_NAME, P02_NAME, P04_NAME  ";
		ls_sql += " FROM  BMSREGT";
		ls_sql += " WHERE  1 = 1 " ;


		if( !StringUtils.isEmpty(wherestring[1]) )
			ls_sql += " AND REG_KIND = '" + wherestring[1]  + "'";

		if( !StringUtils.isEmpty(wherestring[2]) )
			ls_sql += " AND (TMDNUM LIKE '" + wherestring[2] + "%')";

		if( !StringUtils.isEmpty(wherestring[3]) )
			ls_sql += " AND REG_YY = '" + wherestring[3]  + "'";

		if( !StringUtils.isEmpty(wherestring[4]) )
			ls_sql += " AND (REG_NO LIKE '%" + wherestring[4] + "')";

		if( !StringUtils.isEmpty(wherestring[5]) )
			ls_sql += " AND (IO_DEP3_NAME LIKE '%" + wherestring[5] + "%')";

		if( !StringUtils.isEmpty(wherestring[6]) )
			ls_sql += " AND (P01_NAME LIKE '%" + wherestring[6] + "%')";

		if( !StringUtils.isEmpty(wherestring[7]) )
			ls_sql += " AND (P02_NAME LIKE '%" + wherestring[7] + "%')";

		if( !StringUtils.isEmpty(wherestring[8]) )
			ls_sql += " AND (P04_NAME LIKE '%" + wherestring[8] + "%')";

		if( !StringUtils.isEmpty(wherestring[9]) )
			ls_sql += " AND ADDRADR = '" + wherestring[9]  + "'";

		if( !StringUtils.isEmpty(wherestring[10]) )
			ls_sql += " AND SECTION = '" + wherestring[10]  + "'";

		if( !StringUtils.isEmpty(wherestring[11]) )
			ls_sql += " AND (ROAD_NO1 LIKE '%" + wherestring[11] + "')";

		if( !StringUtils.isEmpty(wherestring[12]) )
			ls_sql += " AND (ROAD_NO2 LIKE '%" + wherestring[12] + "')";

		if( !StringUtils.isEmpty(wherestring[13]) )
			ls_sql += " AND IO_DATE >= '" + wherestring[13]  + "'";

		if( !StringUtils.isEmpty(wherestring[14]) )
			ls_sql += " AND IO_DATE <= '" + wherestring[14]  + "'";

		if( !StringUtils.isEmpty(wherestring[15]) )
			ls_sql += " AND (LICENSE_YY LIKE '%" + wherestring[15] + "')";

		if( !StringUtils.isEmpty(wherestring[16]) )
			ls_sql += " AND LICENSE_KIND = '" + wherestring[16]  + "'";

		if( !StringUtils.isEmpty(wherestring[17]) )
			ls_sql += " AND LICENSE_NO1 = '" + wherestring[17]  + "'";

		if( !StringUtils.isEmpty(wherestring[18]) )
			ls_sql += " AND UP_FLOOR_NO <= " + wherestring[18]  ;

		if( !StringUtils.isEmpty(wherestring[19]) )
			ls_sql += " AND UP_FLOOR_NO >= " + wherestring[19]  ;

		if( !StringUtils.isEmpty(wherestring[20]) )
			ls_sql += " AND DN_FLOOR_NO <= " + wherestring[20]  ;

		if( !StringUtils.isEmpty(wherestring[21]) )
			ls_sql += " AND DN_FLOOR_NO >= " + wherestring[21]  ;

		if( !StringUtils.isEmpty(wherestring[22]) )
			ls_sql += " AND PUBLIC_FLAG = '" + wherestring[22]  + "'";

		if( !StringUtils.isEmpty(wherestring[23]) || !StringUtils.isEmpty(wherestring[24])){
			ls_sql += " AND (LAST_RESULT_C = '200' ";
    		if( !StringUtils.isEmpty(wherestring[23])){
				ls_sql += " AND CLOSE_DATE >= '" + wherestring[23]  + "'";
    		}
    			
    		if( !StringUtils.isEmpty(wherestring[24])){
				ls_sql += " AND CLOSE_DATE <= '" + wherestring[24]  + "'";
    		}
			ls_sql += " )";
			
		}

		if( !StringUtils.isEmpty(wherestring[25]) || !StringUtils.isEmpty(wherestring[26])){
			ls_sql += " AND (LAST_RESULT_C = '004' ";
    		if( !StringUtils.isEmpty(wherestring[25])){
				ls_sql += " AND CLOSE_DATE >= '" + wherestring[25]  + "'";
    		}
    			
    		if( !StringUtils.isEmpty(wherestring[26])){
				ls_sql += " AND CLOSE_DATE <= '" + wherestring[26]  + "'";
    		}
			ls_sql += " )";
		}


		if( !StringUtils.isEmpty(wherestring[27]) || !StringUtils.isEmpty(wherestring[28])){
			ls_sql += " AND (LAST_RESULT_C = '005' ";
    		if( !StringUtils.isEmpty(wherestring[27])){
				ls_sql += " AND CLOSE_DATE >= '" + wherestring[27]  + "'";
    		}
    			
    		if( !StringUtils.isEmpty(wherestring[28])){
				ls_sql += " AND CLOSE_DATE <= '" + wherestring[28]  + "'";
    		}
			ls_sql += " )";
		}


		if( !StringUtils.isEmpty(wherestring[29]) )
			ls_sql += " AND WORK_DAYS >= " + wherestring[29] ;

		if( !StringUtils.isEmpty(wherestring[30]) )
			ls_sql += " AND WORK_DAYS <= " + wherestring[30] ;


		if( !StringUtils.isEmpty(wherestring[31]) )
			ls_sql += " AND AREA_FLAG = '" + wherestring[31]  + "'";


		ls_sql += " ORDER BY 1 ";
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

         rds = new String[(int)li_total_row ][19];
		 //double d_tot[] = new double[10];
         //給初始值
         for (int m=0;m<rds.length;m++){
         	for(int n=0;n<rds[m].length;n++){
         		rds[m][n] = "";
         	}
         }

         while( rows1 != null && rows1.hasMoreElements() ){
            DbRow row2 = (DbRow) rows1.nextElement();

             rds[i][0] = Utils.convertToString(row2.get("REG_NO"));
             rds[i][1] = Utils.convertToString(DBTools.dLookUp("code_desc", "BLDCODE", "code_type = 'OFC' AND code_seq ='"+Utils.convertToString(row2.get("REG_KIND")) + "'", "SynctConn"));
            
             rds[i][2] = Utils.convertToString(row2.get("TMDNUM"));
             rds[i][3] = Utils.convertToString(row2.get("LICENSE_NO1"));
             rds[i][4] = Utils.convertToString(row2.get("IO_DATE"));
             rds[i][5] = Utils.convertToString(row2.get("IO_DEP3_NAME"));
             rds[i][6] = Utils.convertToString(row2.get("P01_NAME"));
             rds[i][7] = Utils.convertToString(row2.get("P02_NAME"));
             rds[i][8] = Utils.convertToString(row2.get("P04_NAME"));
             rds[i][9] = Utils.convertToString(row2.get("UP_FLOOR_NO"));
             rds[i][10] = Utils.convertToString(row2.get("DN_FLOOR_NO"));
             
             if( !StringUtils.isEmpty(Utils.convertToString(row2.get("PUBLIC_FLAG"))) && Utils.convertToString(row2.get("PUBLIC_FLAG")).equals("Y")){
                rds[i][11] = "是";
             }
             rds[i][12] = Utils.convertToString(DBTools.dLookUp("code_desc", "BLDCODE", "code_type = 'TMDARE' AND code_seq ='"+Utils.convertToString(row2.get("AREA_FLAG")) + "'", "SynctConn"));
 
             rds[i][13] = Utils.convertToString(DBTools.dLookUp("NAME", "BMSEMP", "EMPNO ='"+Utils.convertToString(row2.get("ACTN_EMPL")) + "'", "SynctConn"));

             rds[i][14] = Utils.convertToString(row2.get("RELATION"));
       
             rds[i][15] = Utils.convertToString(DBTools.dLookUp("code_desc", "BLDCODE", "code_type = 'RES' AND code_seq ='"+Utils.convertToString(row2.get("LAST_RESULT_C")) + "'", "SynctConn"));
             rds[i][16] = Utils.convertToString(row2.get("CLOSE_DATE"));
             rds[i][17] = Utils.convertToString(row2.get("WORK_DAYS"));
             rds[i][18] = Utils.convertToString(row2.get("REAL_WORK_DAYS"));

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
            System.err.println("mc101.java: before getDataValue.");
            data=getDataValue(wherestring);
            System.err.println("mc101.java: end getDataValue.");

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
//System.err.println("mc101.java: before printHeader.");
                //printHeader(wherestring);
//System.err.println("mc101.java: end printHeader.");
	    		//寫入detail
	      		for(int j=0;j<onepage_detail;j++) {
//System.err.println("mc101.java: page =" + page);
                        if(data.length>onepage_detail * page + j){

//System.err.println("mc101.java: pastePageBodyStyleBlock start" );
				    		//pasteMergedRegion(sheet, region, j , 0);
				    		pastePageBodyStyleBlock(sheet, body_style, j  + dtl_start_row, 0);
				    		pastePageBodyValueBlock(sheet, body_value, j  + dtl_start_row, 0);
				    		
				    		
//System.err.println("mc101.java: pastePageBodyStyleBlock end" );
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
         System.err.println("MC101:execOut error is "+e.toString());
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
                    
                    for (int g = 0 ;g <=18 ; g++){
	             	    cell1 = row.getCell((short)(g));
	               	    setBig5CellValue(data1[g],cell1);
                    }


               }catch(Exception e)     {
                  System.err.println("bi30101:printPageBody error is " + e);
               }

    }


 }