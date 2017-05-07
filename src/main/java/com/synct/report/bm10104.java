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

public class bm10104 extends Ole2Adapter {

	private		int onepage_detail = 20000;     //�@�������X�Cdetail
	private    	int dtl_start_row = 3;      //detail�qpage�̪��ĴX�C�}�l
	private    	int dtl_cols = 2;           //detail��Ʀ��X��
	private    	String execlfilename = "bm10104.xls";  //excel�ɦW

    public bm10104() {
        page_rows = 20000;     //�@�������X�C
	}


    //�e������
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

	HSSFCellStyle[][] header_style;   //header�϶���style
	HSSFCellStyle[][] body_style;     //body�϶���style

	String[][] data;
	String[][] header_value;          //header�϶������W��,�μ���
	String[][] body_value;            //body�϶������W��,�μ���

	Region[] region;                  //�X���x�s��}�C


    //�̵e������q��Ʈw���o���
	public String[][] getDataValue(String[] wherestring)throws Exception{


		String s_REG_YY = wherestring[0];
		String s_REG_NO = wherestring[1];


		String ls_sql = "";

		ls_sql += " SELECT SEQ, LMD_MEMO, LMD_VOID ";
		ls_sql += " FROM  LICENSEMEMO_DE";
		ls_sql += " WHERE  LM_SEQ = " + wherestring[0];

        //�Ŀ�
		if( !StringUtils.isEmpty(wherestring[1]) )
			ls_sql += " AND LMD_VOID = " + wherestring[1] ;

		ls_sql += " ORDER BY SEQ ";



       	JDBCConnection conn =  JDBCConnectionFactory.getJDBCConnection("SynctConn");
         System.err.println(ls_sql);

         //����`����
         int li_total_row = 0;

         //������
         Enumeration rows1 = null;
         Enumeration rows2 = null;
         int i = 0;
         rows1 = conn.getRows(ls_sql);
         rows2 = conn.getRows(ls_sql);
         conn.closeConnection();
         //�p���`����
         while( rows2 != null && rows2.hasMoreElements() ){
            DbRow row2 = (DbRow) rows2.nextElement();
            li_total_row++;
         }
         System.err.println("li_total_row="+li_total_row);
         String [][] rds=null;
         rds = new String[(int)li_total_row ][3];
		 //double d_tot[] = new double[10];
         //����l��
         for (int m=0;m<rds.length;m++){
         	for(int n=0;n<rds[m].length;n++){
         		rds[m][n] = "";
         	}
         }

         while( rows1 != null && rows1.hasMoreElements() ){
            DbRow row2 = (DbRow) rows1.nextElement();

            rds[i][0] = Utils.convertToString(i+1);
            rds[i][1]  = Utils.convertToString(row2.get("LMD_MEMO"));
            
			if(!StringUtils.isEmpty(Utils.convertToString(row2.get("LMD_VOID"))) && Utils.convertToString(row2.get("LMD_VOID")).equals("1")){
				rds[i][2]  = "�O";
			}else{
				rds[i][2]  = "";
			}
            

            i ++;
         }


         return rds;
	}

    //��g����
	private void printHeader(String[] wherestring) throws Exception{
		//�g�X����
        HSSFRow pageRow = sheet.getRow(page * page_rows);
        HSSFCell pageCell = pageRow.getCell((short)1);


        pageRow = sheet.getRow(page * page_rows + 1);
        pageCell = pageRow.getCell((short)0);
        setBig5CellValue( "���Ӹ��X�G" + Utils.convertToString(DBTools.dLookUp("LM_LICNUM", "LICENSEMEMO", "SEQ="+wherestring[0], "SynctConn")),pageCell);

	}

    //��g����
	private void printFoot(String printpage) throws Exception{
		   //�g�X����
           HSSFRow pageRow = sheet.getRow((page * page_rows) + 13);
           HSSFCell pageCell = pageRow.getCell((short)0);
		   setBig5CellValue(printpage ,pageCell);

	}


 	/**
  	*<br>�ت��G��X����
  	*<br>�ѼơG �L
  	*<br>�Ǧ^�Gboolean
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

    //����Excel��
    public void execOut(String userid,String[] wherestring) throws Exception{
      FileOutputStream fileOut = null;
      try {
            //�i��Ʈw�d��
            System.err.println("bi30101.java: before getDataValue.");
            data=getDataValue(wherestring);
            System.err.println("bi30101.java: end getDataValue.");

            //�ƻs���Y�˦�
 	        //header_style = copyPageHeaderStyle(sheet, 0,0,dtl_start_row,dtl_cols); // (int start_row, int start_col, int rows ��detail�}�l���header���C��, int cols)

 	        //�ƻs���Y�x�s���
            //header_value = copyPageHeaderValue(sheet, 0,0,dtl_start_row,dtl_cols);

            //�ƻs���˦�
            body_style = copyPageBodyStyleBlock(sheet, 3,0,4,dtl_cols);       // (int row, int start col, int cols)

            //�ƻs���x�s���
            body_value = copyPageBodyValueBlock(sheet, 3,0,4,dtl_cols);

            //�ƻs�Ĥ@���������X���x�s��
            region = copyMergedRegion(sheet);
            //System.err.println("1");
            //���p��X�`����,�j�餤�w��C���B�z��Ȫ��ʧ@,�]�i�H���`���ƥh�]�j��
            int total_page = 0;
            total_page=((data.length - 1)/onepage_detail) + 1;

            //�C���������`�p��
            int total=0;
            int totalCount=0;

    		for(int i=0;i<total_page;i++) {
    			//���K�W���e�ƻs�����  ,�b���s�����[�Jheader
    			if(page != 0) {
					//pastePageHeaderStyle(sheet, header_style, (page_rows * page), 0);
					//pastePageHeaderValue(sheet, header_value, (page_rows * page), 0);
		    		pasteMergedRegion(sheet, region, (page_rows * page), 0);
		    		pastePageBodyStyleBlock(sheet, body_style, (page_rows * page  + dtl_start_row), 0);
		    		pastePageBodyValueBlock(sheet, body_value, (page_rows * page) + dtl_start_row, 0);

		    	}
    			//�񭶭�����
//System.err.println("bi30101.java: before printHeader.");
                printHeader(wherestring);
//System.err.println("bi30101.java: end printHeader.");
	    		//�g�Jdetail
	      		for(int j=0;j<onepage_detail;j++) {
                        if(data.length>onepage_detail * page + j){
				    		pastePageBodyStyleBlock(sheet, body_style, j  + dtl_start_row, 0);
				    		pastePageBodyValueBlock(sheet, body_value, j  + dtl_start_row, 0);

         	    			printPageBody(onepage_detail * page + j + 1,data[onepage_detail * page + j],(page_rows * page) + dtl_start_row + j * 1);
     	    			}else{
            			    break;
     	    			}

	    		}
	    		//�񭶧�����
	    		//printFoot("��" + (page + 1) + "���A�@" + (total_page) + "��");

                //����
	    		setPageBreak(ps);
	    	}
            //��XExcel��
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
         		    //�[�J���
         		    HSSFRow row = sheet.getRow(rowno);
                    HSSFCell cell1 =null;

             	    cell1 = row.getCell((short)(0));
               	    setBig5CellValue(data1[0],cell1);


              	    cell1 = row.getCell((short)(1));
              	    setBig5CellValue(data1[1],cell1);

              	    //cell1 = row.getCell((short)(31));                                    
              	    //setBig5CellValue(data1[2],cell1);


               }catch(Exception e)     {
                  System.err.println("bi30101:printPageBody error is " + e);
               }

    }


 }