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
import com.codecharge.util.*;
import com.codecharge.events.*;
import com.codecharge.db.*;
import com.codecharge.validation.*;

public class bm10101_4 extends Ole2Adapter {

	private		int onepage_detail = 20000;     //一頁報表有幾列detail
	private    	int dtl_start_row = 3;      //detail從page裡的第幾列開始
	private    	int dtl_cols = 2;           //detail資料有幾欄
	private    	String execlfilename = "bm10101_4.xls";  //excel檔名

    public bm10101_4() {
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
	public HSSFSheet sheet,sheet1,sheet2,sheet3;
	public HSSFPrintSetup ps;

	HSSFCellStyle[][] header_style;   //header區塊的style
	HSSFCellStyle[][] body_style;     //body區塊的style

	String[] data;
    List<String> data_other; //= new ArrayList<String>();

	String[][] header_value;          //header區塊的欄位名稱,或標籤
	String[][] body_value;            //body區塊的欄位名稱,或標籤

	Region[] region;                  //合併儲存格陣列


    //依畫面條件從資料庫取得資料
	public String[] getDataValue(String[] wherestring)throws Exception{

		String ls_sql = "";

		ls_sql += " SELECT GETBRKLICS('"+wherestring[0].trim()+"') INFO FROM DUAL ";

       	JDBCConnection conn =  JDBCConnectionFactory.getJDBCConnection("SynctConn");
         System.err.println(ls_sql);
         //抓取資料
         Enumeration rows1 = null;
         int i = 0;
         rows1 = conn.getRows(ls_sql);
         conn.closeConnection();
         String[] rds = new String[60 ];

         String ldata = null;
         while( rows1 != null && rows1.hasMoreElements() ){
            DbRow row2 = (DbRow) rows1.nextElement();
           ldata  = Utils.convertToString(row2.get("INFO"));
            
         }

         rds = ldata.split(";");

         return rds;
	}
  //依畫面條件從資料庫取得 雜項工作物 資料
  public void getData(String[] wherestring)throws Exception{
    String KEY = wherestring[0];
    long l_cnt = Utils.convertToLong(DBTools.dLookUp("COUNT(*) ","BM_P01", " INDEX_KEY ='"+KEY+"' ", "SynctConn"));
    int i_cnt =(int) l_cnt;
    System.out.println("#######bm10101_4 i_cnt:   "+i_cnt);
        JDBCConnection conn =  JDBCConnectionFactory.getJDBCConnection("SynctConn");
         
         //String[] ds = new String[300];
         //String[] ds_new = new String[500];
        List<String> ds = new ArrayList<String>();
        data_other = new ArrayList<String>(); 
        String[] p01  ; //new String[6]
         //抓取資料
         Enumeration rows1 = null;
         DbRow CurrentRecord;

    String ls_sql = " ",temp=" ";
    //ds[0]= "起造人及建物門牌：";
        ds.add("申請人及建物門牌：");
    int x=0;
    for(int n=1;n<=i_cnt;n++){
        ls_sql = " select GETP01X('"+KEY+"',"+Integer.toString(n).trim()+") info from dual  ";  //
       // System.out.println("#######bm10101_4 ls_sql:   "+ls_sql);

       rows1 = conn.getRows(ls_sql);
       
        while (rows1 != null && rows1.hasMoreElements()) {
          CurrentRecord = (DbRow) rows1.nextElement();
          temp = Utils.convertToString(CurrentRecord.get("info"));
        }
        p01=temp.split(";");
       // System.out.println("#######bm10101_4 p01[0]:   "+p01[0]);
        //System.out.println("#######bm10101_4 p01[1]:   "+p01[1]);
        //ds[2*n-1]=p01[0];
        ds.add(p01[0]);
        //ds[2*n]=p01[1];
        ds.add(p01[1]);
       // System.out.println("#######bm10101_4 ds[2*n-1]:   "+ds[2*n-1]);
       // System.out.println("#######bm10101_4 ds[2*n]:   "+ds[2*n]);

    }
    x=2*i_cnt;
    
    l_cnt = Utils.convertToLong(DBTools.dLookUp("COUNT(*) ","BM_LAN", " INDEX_KEY ='"+KEY+"' ", "SynctConn"));
    i_cnt=(int) l_cnt/3;
    x++;
   ds.add("地號表：");

    for(int n=1;n<=i_cnt;n++){
        ls_sql = " select GETLANS('"+KEY+"',"+Integer.toString(n).trim()+") info from dual  ";  //
        System.out.println("#######bm10101_4 ls_sql:   "+ls_sql);

       rows1 = conn.getRows(ls_sql);
       
        while (rows1 != null && rows1.hasMoreElements()) {
          CurrentRecord = (DbRow) rows1.nextElement();
          temp = Utils.convertToString(CurrentRecord.get("info"));
        }
        //ds[x+n]=temp;
        ds.add(temp);
    }

    x+=i_cnt;
    
    /*
    x++;
    //ds[x]= "建築物概要：";
    ds.add("建築物概要：");
    l_cnt = Utils.convertToLong(DBTools.dLookUp("COUNT(*) ","BM_STAIR", " INDEX_KEY ='"+KEY+"' ", "SynctConn"));
    i_cnt = (int)l_cnt;
    for(int n=1;n<=i_cnt;n++){
        ls_sql = " select GETSTAIRS('"+KEY+"',"+Integer.toString(n).trim()+") info from dual  ";  //
        System.out.println("#######bm10101_4 ls_sql:   "+ls_sql);

       rows1 = conn.getRows(ls_sql);
       
        while (rows1 != null && rows1.hasMoreElements()) {
          CurrentRecord = (DbRow) rows1.nextElement();
          temp = Utils.convertToString(CurrentRecord.get("info"));
        }
        //ds[x+n]=temp;
        ds.add(temp);

    }

     x+=i_cnt;
    x++;
    //ds[x]= "雜項工作物：";
    ds.add("雜項工作物：");

    ls_sql += " select  Comb_Work(CONSNAME,BUILDING_KIND,LENGTH,HEIGHT,WIDE,AREA,CONNUM,DESE) other_work FROM BM_WORK ";  //
    ls_sql += " WHERE INDEX_KEY = '"+KEY+"'  order by PERSON_SEQ "; 

    rows1 = conn.getRows(ls_sql);
      while (rows1 != null && rows1.hasMoreElements()) {
        CurrentRecord = (DbRow) rows1.nextElement();
        x++;
       // ds[x] = Utils.convertToString(CurrentRecord.get("other_work"));
        ds.add( Utils.convertToString(CurrentRecord.get("other_work")));
      }

     x++;
     //ds[x]= "停車空間      設置類別     車位分類    檢討類別    室內/外    地上/下   輛數   面積(㎡) ";
     ds.add("停車空間      設置類別     車位分類    檢討類別    室內/外    地上/下   輛數   面積(㎡) ");
   l_cnt = Utils.convertToLong(DBTools.dLookUp("COUNT(*) ","BM_PARK", " INDEX_KEY ='"+KEY+"' ", "SynctConn"));
    i_cnt = (int)l_cnt;
    for(int n=1;n<=i_cnt;n++){
        ls_sql = " select GETPARKS('"+KEY+"',"+Integer.toString(n).trim()+") info from dual  ";  //

       rows1 = conn.getRows(ls_sql);
       
        while (rows1 != null && rows1.hasMoreElements()) {
          CurrentRecord = (DbRow) rows1.nextElement();
          temp = Utils.convertToString(CurrentRecord.get("info"));
        }
        //ds[x+n]=temp;
         ds.add(temp);

    }
     x+=i_cnt;
     */
     x++;
     //ds[x]= "加註事項: ";
     ds.add("加註事項: ");
     x++;
    // ds[x]= "【適用法令概要】";
     ds.add("【適用法令概要】");
    l_cnt = Utils.convertToLong(DBTools.dLookUp("COUNT(*) ","BM_PARK", " INDEX_KEY ='"+KEY+"' ", "SynctConn"));
    i_cnt = (int)l_cnt;
    ls_sql = " select GETLAWS('"+KEY+"') info from dual  ";  //
    rows1 = conn.getRows(ls_sql);
        while (rows1 != null && rows1.hasMoreElements()) {
          CurrentRecord = (DbRow) rows1.nextElement();
          temp = Utils.convertToString(CurrentRecord.get("info"));
        }

        p01=temp.split(";");
        for (String token:p01) {
             x++;
            //ds[x]= token;
           if(!StringUtils.isEmpty(token))
           {ds.add(token);}             
            
         }

    l_cnt = Utils.convertToLong(DBTools.dLookUp("COUNT(*) ","BM_MEMO", " INDEX_KEY ='"+KEY+"' ", "SynctConn"));
    i_cnt = (int)l_cnt;
    for(int n=1;n<=i_cnt;n++){
        ls_sql = " select GETMEMOS('"+KEY+"',"+Integer.toString(n).trim()+") info from dual  ";  //
        System.out.println("#######bm10101_4 ls_sql:   "+ls_sql);

       rows1 = conn.getRows(ls_sql);
       
        while (rows1 != null && rows1.hasMoreElements()) {
          CurrentRecord = (DbRow) rows1.nextElement();
          temp = Utils.convertToString(CurrentRecord.get("info"));
        }
        p01=temp.split(";");
        
         for (String token:p01) {
             x++;
           // ds[x]= token;
            ds.add(token);
            System.out.println("#######bm10101_4 token:   "+token);
         }
    }
      conn.closeConnection();
    //System.out.println("data length   total_page: "+total_page );
    //System.out.println(" ds[x] :x="+x+"  ,ds[x]= "+ds[x] );
    //int pages=((ds.length - 1)/51) + 1; - 1 
    int pages=((ds.size() )/51) + 1;  
    int pg = 0;
    System.out.println("ds size :"+ds.size()+"  pages: "+pages );
    for(int n=1;n<= 55*(pages+1);n++){

        if (n<=55) {pg = 0 ;}
        if (n>=56 && n<=110) {pg = 1;}
        if (n>=111 && n<=165) {pg = 2;}
        if (n>=166 && n<=220) {pg = 3;}
        if (n>=221 && n<=275) {pg = 4;}
        if (n>=276 && n<=330) {pg = 5;}
       
       if (n==(55*pg+1 )) 
       {
        // data_other[n]="新北市政府工務局     使用執照附表";
        data_other.add("新北市政府工務局     使用執照附表       ");
       }
      if ( (n==55*(pg+1 )) || (  n==(55*(pg+1) -1))   )  //&& (n<= (ds.size()-1) ) 
       {
        //ds_new[n]="       ";
        data_other.add("       ");
       }
      if (  n == (55*(pg+1) -2)  )  //&& (n<= (ds.size()-1) )
       {
        //ds_new[n]="本附表共"+Integer.toString(pages)+"頁(第 "+Integer.toString(pg+1)+"  頁)";
        data_other.add("本附表共"+Integer.toString(pages)+"頁(第 "+Integer.toString(pg+1)+"  頁)");
       }
      if ( (n>=55*pg+2 ) && (n<=55*(pg+1) -3)   )  //&& (n<= (ds.size()-1) )
       {
        //ds_new[n]= ds[n-2-4*pg];
        int f = n-2-4*pg;
        if (f< ds.size())
        {
            data_other.add(ds.get(n-2-4*pg));
        }
         //System.out.println("ds_new n :"+n+"  ds[n-2-4*pg] "+ds[n-2-4*pg] );
       }

    }  

      System.out.println("data_other size :"+data_other.size() ); 

     // return ds_new; 
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
            System.err.println("bm10101_4.java: before getDataValue.");
            data=getDataValue(wherestring);
            getData(wherestring);
            System.err.println("bm10101_4.java: end getDataValue.");


            //複製表身樣式
            body_style = copyPageBodyStyleBlock(sheet, 0,0,45,45);       // (int row, int start col, int cols)

            //複製表身儲存格值
            body_value = copyPageBodyValueBlock(sheet, 0,0,45,45);

            //複製第一頁報表內的合併儲存格
            //region = copyMergedRegion(sheet);
            //System.err.println("1");
            //先計算出總頁數,迴圈中針對每頁處理塞值的動作,也可以用總筆數去跑迴圈
        	pastePageBodyStyleBlock(sheet, body_style, 0, 0);
		    pastePageBodyValueBlock(sheet, body_value, 0, 0);
    			//填頁首的值
	    		//寫入detail
     	    printPageBody(0,data,0);

            //複製第二頁
            body_style = copyPageBodyStyleBlock(sheet1, 0,0,55,35);       // (int row, int start col, int cols)

            //複製表身儲存格值
            body_value = copyPageBodyValueBlock(sheet1, 0,0,55,35);

            //複製第二頁報表內的合併儲存格
            region = copyMergedRegion(sheet1);

            //pasteMergedRegion(sheet1, region, 0, 0);
            //先計算出總頁數,迴圈中針對每頁處理塞值的動作,也可以用總筆數去跑迴圈
            int total_page = 0;
            //每頁頁尾的總計值
            int total=0;
            int totalCount=0;

            total_page=((data_other.size() - 1)/55) + 1;
            System.out.println("data length   total_page: "+total_page );
            for(int i=0;i<total_page;i++) {  
                pasteMergedRegion(sheet1, region, 0+55*i , 0);    
                pastePageBodyStyleBlock(sheet1, body_style, 0+55*i, 0);
                pastePageBodyValueBlock(sheet1, body_value, 0+55*i, 0);
            }
            //填格子
            printPageBody1(0,data,0); 

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
            printPageBody2(0,data,0); 
	    		//填頁尾的值
            //複製第四頁
            body_style = copyPageBodyStyleBlock(sheet3, 0,0,55,35);       // (int row, int start col, int cols)

            //複製表身儲存格值
            body_value = copyPageBodyValueBlock(sheet3, 0,0,55,35);

            //複製第四頁報表內的合併儲存格
            region = copyMergedRegion(sheet3);

            total_page=((data_other.size() - 1)/55) + 1;
            System.out.println("data length   total_page: "+total_page );
            for(int i=0;i<total_page;i++) {  
                pasteMergedRegion(sheet3, region, 0+55*i , 0);    
                pastePageBodyStyleBlock(sheet3, body_style, 0+55*i, 0);
                pastePageBodyValueBlock(sheet3, body_value, 0+55*i, 0);
            }
                printPageBody3(0,data,0); 
	    	
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
        setBig5CellValue(data1[0],cell1);  //U1 LICENSE_DESC

        row = sheet.getRow(1);
        cell1 =row.getCell((short)(6));   
        setBig5CellValue(data1[1],cell1); //G2 P01_NAME


        row = sheet.getRow(2);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[2],cell1);  //G3 ADDR

        row = sheet.getRow(3);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[3],cell1);  //G4 LANNO

        row = sheet.getRow(4);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[4],cell1);  //G5 ADDR

        row = sheet.getRow(5);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[5],cell1);  //G6 USE_CATEGORY_CODE_DESC

        row = sheet.getRow(6);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[6],cell1);  //G7 CHWANG DONG
 
        row = sheet.getRow(10);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[7],cell1);  //G8 BUILDING_CATEGORY


        //*****************************************************************************
        /*
        0 1 2 3 4 5 6 7 8 910111213141516171819202122232425 6 7 8 930 1 2 3 4 5 
        A B C D E F G H I J K L M N O P Q R S T U V W X Y ZAAABACADAEAFAGAHAIAJ        
        */
        row = sheet.getRow(8);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[8],cell1);  //G9 STC1     i

        row = sheet.getRow(9);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[9],cell1);  //G10 AREA1
        
        row = sheet.getRow(10);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[10],cell1);  //G11 HEIGHT1

        row = sheet.getRow(8);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[11],cell1);  //L9 STC1  ii

        row = sheet.getRow(9);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[12],cell1);  //L10 AREA1
        
        row = sheet.getRow(10);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[13],cell1);  //L11 HEIGHT1

        row = sheet.getRow(8);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[14],cell1);  //Q9 STC1 iii

        row = sheet.getRow(9);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[15],cell1);  //Q10 AREA1
        
        row = sheet.getRow(10);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[16],cell1);  //Q11 HEIGHT1

        row = sheet.getRow(8);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[17],cell1);  //V9 STC1 iv

        row = sheet.getRow(9);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[18],cell1);  //V10 AREA1
        
        row = sheet.getRow(10);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[19],cell1);  //V11 HEIGHT1

        row = sheet.getRow(8);
        cell1 =row.getCell((short)(26));
        setBig5CellValue(data1[20],cell1);  //AA9 STC1 v

        row = sheet.getRow(9);
        cell1 =row.getCell((short)(26));
        setBig5CellValue(data1[21],cell1);  //AA10 AREA1
        
        row = sheet.getRow(10);
        cell1 =row.getCell((short)(26));
        setBig5CellValue(data1[22],cell1);  //AA11 HEIGHT1

        row = sheet.getRow(8);
        cell1 =row.getCell((short)(31));
        setBig5CellValue(data1[23],cell1);  //AF9 STC1 vi

        row = sheet.getRow(9);
        cell1 =row.getCell((short)(31));
        setBig5CellValue(data1[24],cell1);  //AF10 AREA1
        
        row = sheet.getRow(10);
        cell1 =row.getCell((short)(31));
        setBig5CellValue(data1[25],cell1);  //AF11 HEIGHT1



        row = sheet.getRow(11);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[26],cell1);  //G12 STC1 vii

        row = sheet.getRow(12);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[27],cell1);  //G13 AREA1
        
        row = sheet.getRow(13);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[28],cell1);  //G14 HEIGHT1

        row = sheet.getRow(11);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[29],cell1);  //L12 STC1 viii

        row = sheet.getRow(12);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[30],cell1);  //L13 AREA1
        
        row = sheet.getRow(13);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[31],cell1);  //L14 HEIGHT1

        row = sheet.getRow(11);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[32],cell1);  //Q12 STC1 ix

        row = sheet.getRow(12);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[33],cell1);  //Q13 AREA1
        
        row = sheet.getRow(13);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[34],cell1);  //Q14 HEIGHT1

        row = sheet.getRow(11);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[35],cell1);  //V12 STC1 x

        row = sheet.getRow(12);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[36],cell1);  //V13 AREA1
        
        row = sheet.getRow(13);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[37],cell1);  //V14 HEIGHT1

        row = sheet.getRow(14);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[38],cell1);  //G15 OTHERS_NAME

	    row = sheet.getRow(15);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[39],cell1);  //G16 USE LICE_DESC

        row = sheet.getRow(16);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[40],cell1);  //G17 PRICE


      
        row = sheet.getRow(17);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[41],cell1);  //G18 APPROVE_LICE_DATE

        row = sheet.getRow(17);
        cell1 =row.getCell((short)(26));
        setBig5CellValue(data1[42],cell1);  //AA18 RECEIVE_LICE_DATE

        row = sheet.getRow(20);
        cell1 =row.getCell((short)(1));
        setBig5CellValue("上給    "+data1[1],cell1);  //B21 P01_NAME
        

        row = sheet.getRow(27);
        cell1 =row.getCell((short)(0));
        setBig5CellValue(data1[43],cell1);  //A28 IDENTIFY_LICE_DATE


        }catch(Exception e)     {
          System.err.println("bm10101_4:printPageBody error is " + e);
       }

    }

    public void printPageBody1(int k,String[] data1, int rowno) throws IOException {
        try{

        HSSFRow row = sheet1.getRow(0);
        HSSFCell  cell1 =row.getCell((short)(0));
        int j=0;
        
        for(int i=0;i<data_other.size() ;i++) {  
          row = sheet1.getRow(i); //1+i
          cell1 =row.getCell((short)(0));
          //setBig5CellValue(data_other[i],cell1); 
          setBig5CellValue(data_other.get(i),cell1); 
          j=i;
              if (i>220) {
                 System.out.println("bm10101_4 sheet3:data_other.get(i),i= " + i+","+data_other.get(i));
              }
          }
        row = sheet1.getRow(3+j);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("以下空白",cell1); //        

      }catch(Exception e)     {
          System.err.println("bm10101_4:printPageBody1 error is " + e);
      }
  }

public void printPageBody2(int k,String[] data1, int rowno) throws IOException {
       try{
        //加入資料
        HSSFRow row = sheet2.getRow(0);
        HSSFCell cell1 =row.getCell((short)(20));
        setBig5CellValue(data1[0],cell1);  //U1 LICENSE_DESC

        row = sheet2.getRow(1);
        cell1 =row.getCell((short)(6));   
        setBig5CellValue(data1[1],cell1); //G2 P01_NAME


        row = sheet2.getRow(2);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[2],cell1);  //G3 ADDR

        row = sheet2.getRow(3);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[3],cell1);  //G4 LANNO

        row = sheet2.getRow(4);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[4],cell1);  //G5 ADDR

        row = sheet2.getRow(5);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[5],cell1);  //G6 USE_CATEGORY_CODE_DESC

        row = sheet2.getRow(6);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[6],cell1);  //G7 CHWANG DONG
 
        row = sheet2.getRow(10);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[7],cell1);  //G8 BUILDING_CATEGORY


        //*****************************************************************************
        /*
        0 1 2 3 4 5 6 7 8 910111213141516171819202122232425 6 7 8 930 1 2 3 4 5 
        A B C D E F G H I J K L M N O P Q R S T U V W X Y ZAAABACADAEAFAGAHAIAJ        
        */
        row = sheet2.getRow(8);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[8],cell1);  //G9 STC1

        row = sheet2.getRow(9);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[9],cell1);  //G10 AREA1
        
        row = sheet2.getRow(10);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[10],cell1);  //G11 HEIGHT1

        row = sheet2.getRow(8);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[11],cell1);  //L9 STC1

        row = sheet2.getRow(9);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[12],cell1);  //L10 AREA1
        
        row = sheet2.getRow(10);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[13],cell1);  //L11 HEIGHT1

        row = sheet2.getRow(8);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[14],cell1);  //Q9 STC1

        row = sheet2.getRow(9);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[15],cell1);  //Q10 AREA1
        
        row = sheet2.getRow(10);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[16],cell1);  //Q11 HEIGHT1

        row = sheet2.getRow(8);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[17],cell1);  //V9 STC1

        row = sheet2.getRow(9);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[18],cell1);  //V10 AREA1
        
        row = sheet2.getRow(10);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[19],cell1);  //V11 HEIGHT1

        row = sheet2.getRow(8);
        cell1 =row.getCell((short)(26));
        setBig5CellValue(data1[20],cell1);  //AA9 STC1

        row = sheet2.getRow(9);
        cell1 =row.getCell((short)(26));
        setBig5CellValue(data1[21],cell1);  //AA10 AREA1
        
        row = sheet2.getRow(10);
        cell1 =row.getCell((short)(26));
        setBig5CellValue(data1[22],cell1);  //AA11 HEIGHT1

        row = sheet2.getRow(8);
        cell1 =row.getCell((short)(31));
        setBig5CellValue(data1[23],cell1);  //AF9 STC1

        row = sheet2.getRow(9);
        cell1 =row.getCell((short)(31));
        setBig5CellValue(data1[24],cell1);  //AF10 AREA1
        
        row = sheet2.getRow(10);
        cell1 =row.getCell((short)(31));
        setBig5CellValue(data1[25],cell1);  //AF11 HEIGHT1



        row = sheet2.getRow(11);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[26],cell1);  //G12 STC1

        row = sheet2.getRow(12);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[27],cell1);  //G13 AREA1
        
        row = sheet2.getRow(13);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[28],cell1);  //G14 HEIGHT1

        row = sheet2.getRow(11);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[29],cell1);  //L12 STC1

        row = sheet2.getRow(12);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[30],cell1);  //L13 AREA1
        
        row = sheet2.getRow(13);
        cell1 =row.getCell((short)(11));
        setBig5CellValue(data1[31],cell1);  //L14 HEIGHT1

        row = sheet2.getRow(11);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[32],cell1);  //Q12 STC1

        row = sheet2.getRow(12);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[33],cell1);  //Q13 AREA1
        
        row = sheet2.getRow(13);
        cell1 =row.getCell((short)(16));
        setBig5CellValue(data1[34],cell1);  //Q14 HEIGHT1

        row = sheet2.getRow(11);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[35],cell1);  //V12 STC1

        row = sheet2.getRow(12);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[36],cell1);  //V13 AREA1
        
        row = sheet2.getRow(13);
        cell1 =row.getCell((short)(21));
        setBig5CellValue(data1[37],cell1);  //V14 HEIGHT1

        row = sheet2.getRow(14);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[38],cell1);  //G15 OTHERS_NAME

	    row = sheet2.getRow(15);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[39],cell1);  //G16 USE LICE_DESC

        row = sheet2.getRow(16);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[40],cell1);  //G17 PRICE


      
        row = sheet2.getRow(17);
        cell1 =row.getCell((short)(6));
        setBig5CellValue(data1[41],cell1);  //G18 APPROVE_LICE_DATE

        row = sheet2.getRow(17);
        cell1 =row.getCell((short)(26));
        setBig5CellValue(data1[42],cell1);  //AA18 RECEIVE_LICE_DATE

        row = sheet2.getRow(20);
        cell1 =row.getCell((short)(1));
        setBig5CellValue("上給    "+data1[1],cell1);  //B21 P01_NAME
        

        row = sheet2.getRow(27);
        cell1 =row.getCell((short)(0));
        setBig5CellValue(data1[43],cell1);  //A28 IDENTIFY_LICE_DATE

       }catch(Exception e)     {
          System.err.println(" sheet2:printPageBody error is " + e);
       }

    }

    public void printPageBody3(int k,String[] data1, int rowno) throws IOException {
  		try{

       HSSFRow row = sheet3.getRow(0);
        HSSFCell  cell1 =row.getCell((short)(0));
        int j=0;
        for(int i=0;i<=data_other.size() ;i++) {  //A3 A4 A5 A6 A7   55*5
          row = sheet3.getRow(i); //1+i
          cell1 =row.getCell((short)(0));
          setBig5CellValue(data_other.get(i),cell1); 
          j=i;
              if (i>220) {
                 System.out.println("bm10101_4 sheet3:data_other.get(i),i= " + i+","+data_other.get(i));
              }
          }

        row = sheet3.getRow(3+j);
        cell1 =row.getCell((short)(0));
        setBig5CellValue("以下空白",cell1); //        

      }catch(Exception e)     {
          System.err.println("bm10101_4:printPageBody3 error is " + e);
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

