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

public class cm40101 extends Ole2Adapter {

	private		int onepage_detail = 20000;     //
	private    	int dtl_start_row = 3;      //
	private    	int dtl_cols = 2;           //
	private    	String execlfilename = "cm40101.xls";  //

    public cm40101() {
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
	public HSSFSheet sheet,sheet2,sheet3;
 
	public HSSFPrintSetup ps;

	HSSFCellStyle[][] header_style;   //
	HSSFCellStyle[][] body_style;     //

	ArrayList<String> data = new ArrayList<String>();
	String[][] header_value;          //
	String[][] body_value;            //

	Region[] region;                  //


    //
	public void getDataValue(String[] wherestring, Integer sheetNo)throws Exception{


		String s_REG_YY = wherestring[0];
		String s_REG_NO = wherestring[1];
    String[] rds = new String[100 ];

		String ls_sql = "";

    if (sheetNo == 1) 
		ls_sql = " SELECT *  FROM  guild WHERE  guild_id = 1 ";

    if (sheetNo == 2) 
    ls_sql = " SELECT *  FROM  guild_1 WHERE  guild_id = 1 ";
		
    if (sheetNo == 3) 
    ls_sql = " SELECT *  FROM  guild_2 WHERE  guild_id = 1 ";

    System.out.println("ls_sql="+ls_sql);
    data.clear();

        //

       	JDBCConnection conn =  JDBCConnectionFactory.getJDBCConnection("SynctConn");
         System.err.println(ls_sql);

         //
         int li_total_row = 0;

         //
         //Enumeration rows1 = null;
         //Enumeration rows2 = null;
         //int i = 0;
          Statement  stmt = conn.createStatement(); //ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY         
          ResultSet rs = stmt.executeQuery(ls_sql);
          ResultSetMetaData metadata = rs.getMetaData();
          int columnCount = metadata.getColumnCount();             
         //rows2 = conn.getRows(ls_sql);
         //
          while (rs.next()) {
              String row = "";
              for (int i = 1; i <= columnCount; i++) {
                  row += rs.getString(i) + ", ";   
                  data.add(rs.getString(i)) ;
              }
              System.out.println("data output:");
              System.out.println(row);
  
          }

            
      
  
          conn.closeConnection();
      

         //return rds;
	}

    //
	private void printHeader(String[] wherestring) throws Exception{
		//
        HSSFRow pageRow = sheet.getRow(page * page_rows);
        HSSFCell pageCell = pageRow.getCell((short)1);
        pageRow = sheet.getRow(page * page_rows + 1);
        pageCell = pageRow.getCell((short)0);
      //  setBig5CellValue( "°õ·Ó¸¹½X¡G" + Utils.convertToString(DBTools.dLookUp("LM_LICNUM", "LICENSEMEMO", "SEQ="+wherestring[0], "SynctConn")),pageCell);

	}

    //
	private void printFoot(String printpage) throws Exception{
		   //
           HSSFRow pageRow = sheet.getRow((page * page_rows) + 13);
           HSSFCell pageCell = pageRow.getCell((short)0);
		 //  setBig5CellValue(printpage ,pageCell);

	}


 	/**
  	*<br>
  	*<br>
  	*<br>
  	*/
 	public synchronized  boolean outXLS(String userid,String[] wherestring) throws Exception{
    try{
        separator =  System.getProperty("file.separator");
	    	fs = new POIFSFileSystem(new FileInputStream(getPath() + "template" + separator + execlfilename));
	    	wb = new HSSFWorkbook(fs);
	    	sheet = wb.getSheetAt(0);
        sheet2 = wb.getSheetAt(1);
        sheet3 = wb.getSheetAt(2);
	    	ps = sheet.getPrintSetup();
	    	sheet.setAutobreaks(false);
        
			  execOut(userid,wherestring); //
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
            System.err.println("before getDataValue:");
            getDataValue(wherestring,1); //data=
            printPageBody();
            getDataValue(wherestring,3); //data=
            printPageBody2();
            getDataValue(wherestring,2); //data=
            printPageBody3();
            System.err.println("end getDataValue:");

	        fileOut = new FileOutputStream(getPath() + "output" + separator + userid + execlfilename);
		      wb.write(fileOut);
      }catch(Exception e) {
         System.err.println("AP30000:execOut error is "+e.toString());
         throw new Exception(e.getMessage());
      }finally{
          fileOut.close();
      }
    }


  public void printPageBody() throws IOException {
    try
    {
	    //HSSFRow row = sheet.getRow(rowno);
        //HSSFCell cell1 =null;
        System.out.println("Start to write excel");
        //System.out.println(data.get(0));
        System.out.println(data.get(1));
        System.out.println(data.get(2));
        System.out.println(data.get(3));

        CellReference cellReference = new CellReference("G3"); 
        HSSFRow row = sheet.getRow(cellReference.getRow());
        HSSFCell cell = row.getCell(cellReference.getCol()); //¤¤¤å  
        setBig5CellValue(data.get(1),cell);
        //cell.setCellValue( data.get(1) );


        cellReference = new CellReference("G63"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(1),cell);

        cellReference = new CellReference("G123"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(1),cell);

        cellReference = new CellReference("G183"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(1),cell);

        cellReference = new CellReference("V3"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("V63"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("V123"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("V183"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AM3"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AM63"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AM123"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AM183"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F4"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F64"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F124"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F184"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);


        cellReference = new CellReference("AM4"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AM64"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AM124"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AM184"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);





        cellReference = new CellReference("U6"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("U66"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("U126"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("U186"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        //cell.setCellValue( "\u7f85\u5ef7\u4e2d\u5e25\u54e5" );
        cellReference = new CellReference("U8"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(5) );
        
        cellReference = new CellReference("U68"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(5) );

        cellReference = new CellReference("U128"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(5) );

        cellReference = new CellReference("U188"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(5) );

        cellReference = new CellReference("U10"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(3) );

        cellReference = new CellReference("U70"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(3) );

        cellReference = new CellReference("U130"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(3) );

        cellReference = new CellReference("U190"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(3) );


        cellReference = new CellReference("U12"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U72"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U132"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U192"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U14"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U74"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U134"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U194"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U14"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U74"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U134"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U194"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(4) );

        cellReference = new CellReference("U16"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(6) );

        cellReference = new CellReference("U76"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(6) );

        cellReference = new CellReference("U136"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(6) );

        cellReference = new CellReference("U196"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(6) );

        cellReference = new CellReference("U18"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U78"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U138"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U198"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );


        cellReference = new CellReference("U20"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U80"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U140"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U200"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U22"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U82"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U142"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U202"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U22"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U82"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U142"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U202"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U26"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U86"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U146"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U206"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

       cellReference = new CellReference("U30"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U90"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U150"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("U210"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        cell.setCellValue( data.get(8) );

        cellReference = new CellReference("AU6"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU66"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU126"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU186"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU7"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU67"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU127"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU187"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AJ8"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AJ68"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AJ128"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AJ188"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        

        cellReference = new CellReference("AU8"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU68"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU128"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU188"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU9"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU69"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU129"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU189"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AU10"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU70"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU130"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU190"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AG14"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG74"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG134"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG194"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        
        cellReference = new CellReference("AG16"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG76"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG136"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG196"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        

        cellReference = new CellReference("AG18"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG78"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG138"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG198"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AG20"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG80"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG140"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG200"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AL22"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL82"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL142"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL202"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AL25"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL85"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL145"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL205"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AL28"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL88"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL148"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL208"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AL31"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL91"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL151"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL211"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("B34"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B94"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B154"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B214"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("M34"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("M94"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("M154"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("M214"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AD34"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD94"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD154"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD214"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("H35"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("H95"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("H155"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("H215"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("B36"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B96"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B156"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B216"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("G36"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("G96"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("G156"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("G216"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("B37"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B97"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B157"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B217"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
 
        cellReference = new CellReference("AN37"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AN97"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AN157"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AN217"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AS37"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AS97"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AS157"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AS217"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AR40"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AR100"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AR160"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AR220"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AU41"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU101"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU161"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU221"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("B43"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B103"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B163"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B223"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        

        cellReference = new CellReference("B44"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B104"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B164"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B224"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("F44"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F104"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F164"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F224"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("H45"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("H105"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("H165"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("H225"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("Y45"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Y105"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Y165"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Y225"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("F46"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F106"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F166"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F226"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        



        cellReference = new CellReference("B47"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B107"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B167"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("B227"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AN47"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AN107"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AN167"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AN227"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AS47"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AS107"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AS167"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AS227"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AR48"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AR108"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AR168"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AR228"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
        cellReference = new CellReference("AU49"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU109"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU169"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU229"); 
        row = sheet.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
        
   	    //cell1 = row.getCell((short)(0));

   	    //setBig5CellValue(data1[0],cell1);
   	    //cell1 = row.getCell((short)(1));
   	    //setBig5CellValue(data1[1],cell1);

      	//cell1 = row.getCell((short)(31));                                    
      	//setBig5CellValue(data1[2],cell1);


    }catch(Exception e){
          System.err.println("cm40101:printPageBody error is " + e);
    }

  }

  public void printPageBody2() throws IOException {
    try
    {
        CellReference cellReference = new CellReference("J3"); 
        HSSFRow row = sheet2.getRow(cellReference.getRow());
        HSSFCell cell = row.getCell(cellReference.getCol()); //¤¤¤å  
        setBig5CellValue(data.get(1),cell);
        //cell.setCellValue( data.get(1) );


        cellReference = new CellReference("J43"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(1),cell);

        cellReference = new CellReference("Q3"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q43"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);
      
        cellReference = new CellReference("W3"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("W43"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG3"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AG43"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL3"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AL43"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AS3"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AS43"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BA3"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BA43"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J4"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J44"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AT4"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AT44"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BD4"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BD44"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AT5"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AT45"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J6"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J46"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J7"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J47"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU6"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU46"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AY6"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AY46"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BH6"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BH46"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU7"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AU47"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AY7"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AY47"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BH7"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BH47"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J8"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J48"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("V8"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("V48"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AT8"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AT48"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J10"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J50"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);


        cellReference = new CellReference("S10"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("S50"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("S10"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("S50"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AB10"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AB50"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AK10"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AK50"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AT10"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AT50"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J11"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J51"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AB11"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AB51"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AK11"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AK51"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AQ11"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AQ51"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("D14"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("D54"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J14"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("J54"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("P14"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("P54"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("V14"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("V54"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AB14"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AB54"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AH14"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AH54"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AN14"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AN54"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BB14"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BB54"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("W15"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("W55"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("W16"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("W56"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE17"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE57"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE18"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE58"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE19"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE59"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE20"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE60"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE21"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE61"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE22"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE62"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE23"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE63"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE24"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE64"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("A27"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BE67"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AW33"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AW73"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BB33"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BB73"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BH33"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("BH73"); 
        row = sheet2.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);




    }catch(Exception e){
          System.err.println("cm40101:printPageBody2 error is " + e);
    }

  }


  public void printPageBody3() throws IOException {
    try
    {
        CellReference cellReference = new CellReference("E3"); 
        HSSFRow row = sheet3.getRow(cellReference.getRow());
        HSSFCell cell = row.getCell(cellReference.getCol()); //¤¤¤å  
        setBig5CellValue(data.get(1),cell);


        cellReference = new CellReference("E43"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("E83"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("E123"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("I3"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("I43"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("I83"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("I123"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("T3"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("T43"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("T83"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("T123"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("X3"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("X43"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("X83"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("X123"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AE3"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AE43"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AE83"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AE123"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F5"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F45"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F85"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F125"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F6"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F46"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F86"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F126"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F7"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F47"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F87"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F127"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F8"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F48"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F88"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F128"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F9"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F49"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F89"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F129"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F10"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F50"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F90"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F130"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F11"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F51"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F91"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F131"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F12"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F52"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F92"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F132"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F13"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F53"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F93"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F133"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F14"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F54"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F94"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F134"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F15"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F55"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F95"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F135"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F16"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F56"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F96"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F136"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F17"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F57"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F97"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F137"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F18"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F58"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F98"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F138"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F19"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F59"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F99"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F139"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F20"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F60"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F100"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F140"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F21"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F61"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F101"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F141"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F22"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F62"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F102"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F142"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F23"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F63"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F103"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("F143"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q6"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q46"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q86"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q126"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q7"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q8"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q9"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q10"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD6"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD7"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD8"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD9"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD10"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q12"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q13"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q14"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q15"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q16"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q17"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD12"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD13"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD14"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD15"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD16"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD17"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q19"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q20"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q21"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q22"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q23"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD19"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD20"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD21"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD22"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AD23"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("A25"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q25"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("W25"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Q26"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("W26"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("T27"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("T31"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("W32"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("Z32"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);

        cellReference = new CellReference("AE32"); 
        row = sheet3.getRow(cellReference.getRow());
        cell = row.getCell(cellReference.getCol()); 
        setBig5CellValue(data.get(2),cell);



    }catch(Exception e){
          System.err.println("cm40101:printPageBody3 error is " + e);
    }

  }




}