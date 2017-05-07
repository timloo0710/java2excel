package com.synct.report;
import java.text.SimpleDateFormat;
import java.util.Date;
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

public class bm20113_2 extends Ole2Adapter {

	private		int onepage_detail = 20000;     //�@�������X�Cdetail
	private    	int dtl_start_row = 2;      //detail�qpage�̪��ĴX�C�}�l
	private    	int dtl_cols = 18;           //detail��Ʀ��X��
	private    	String execlfilename = "bm20113_2.xls";  //excel�ɦW

    public bm20113_2() {
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


		String STEP = "";
	    if(!StringUtils.isEmpty(wherestring[3]) && wherestring[3].equals("3")){
	   	   STEP = "'2','3'";
	    }else{
	   	   STEP = "'" + wherestring[3] + "'";
	    }

		String ls_sql = "";

		ls_sql += " SELECT  CASEID , ITEMID, INFORM, CANCEL, APP_LIC_DATE, RCV_LIC_DATE, COMPLETE_DATE, DELAY, HIGHLIGHT, ";
		ls_sql += " HLTYPE, LICENSE_DESC , DIST, OWNER, O_TEL, CONTRATOR, C_TEL, CONTACT, CC_TEL, SITE_DIRECTOR, ";
		ls_sql += " SD_TEL, ADDRESS, SUPERVISOR, S_TEL, AGENT, A_TEL, DIST_CODE, HIGH_CODE  ";
		ls_sql += " FROM  MBCASEAPP_D";
	    ls_sql += " WHERE CASEID = '" + wherestring[0] + "' AND LICENSE_DESC NOT IN (SELECT LICENSE_DESC FROM MBSTEP WHERE CASEID = '" + wherestring[0] + "' AND STEP IN (" + STEP + ")) ";


		if( !StringUtils.isEmpty(wherestring[1]) )
			ls_sql += " AND DIST_CODE = '" + wherestring[1] + "'";

		if( !StringUtils.isEmpty(wherestring[2]) && wherestring[2].equals("1"))
			ls_sql += " AND HIGH_CODE = '1'";
		else if( !StringUtils.isEmpty(wherestring[2]) && wherestring[2].equals("4"))
			ls_sql += " AND HIGH_CODE <> '1'";

		ls_sql += " ORDER BY DIST, HIGHLIGHT, LICENSE_DESC ";




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
         //System.err.println("li_total_row="+li_total_row);
         String [][] rds=null;
         rds = new String[(int)li_total_row ][dtl_cols];
		 //double d_tot[] = new double[10];
         //����l��
         for (int m=0;m<rds.length;m++){
         	for(int n=0;n<rds[m].length;n++){
         		rds[m][n] = "";
         	}
         }

         while( rows1 != null && rows1.hasMoreElements() ){
            DbRow row2 = (DbRow) rows1.nextElement();
            rds[i][0]  = Utils.convertToString(row2.get("APP_LIC_DATE"));
            rds[i][1]  = Utils.convertToString(row2.get("RCV_LIC_DATE"));
            rds[i][2]  = Utils.convertToString(row2.get("COMPLETE_DATE"));
            rds[i][3]  = Utils.convertToString(row2.get("LICENSE_DESC"));
            
            rds[i][4]  = Utils.convertToString(row2.get("DIST"));
            rds[i][5]   = Utils.convertToString(row2.get("OWNER"));
            rds[i][6]   = Utils.convertToString(row2.get("O_TEL"));
            rds[i][7]   = Utils.convertToString(row2.get("CONTRATOR"));
            rds[i][8]   = Utils.convertToString(row2.get("C_TEL"));
            rds[i][9]  = Utils.convertToString(row2.get("CONTACT"));
            rds[i][10]  = Utils.convertToString(row2.get("CC_TEL"));
            rds[i][11]  = Utils.convertToString(row2.get("SITE_DIRECTOR"));
            rds[i][12]  = Utils.convertToString(row2.get("SD_TEL"));
            rds[i][13]  = Utils.convertToString(row2.get("ADDRESS"));
            rds[i][14]  = Utils.convertToString(row2.get("SUPERVISOR"));
            rds[i][15]  = Utils.convertToString(row2.get("S_TEL"));
            rds[i][16]  = Utils.convertToString(row2.get("AGENT"));
            rds[i][17]  = Utils.convertToString(row2.get("A_TEL"));
 

            i ++;
         }


         return rds;
	}

    //��g����
	private void printHeader(String[] wherestring) throws Exception{
		//�g�X����
        HSSFRow pageRow = sheet.getRow(page * page_rows);
        HSSFCell pageCell = pageRow.getCell((short)1);

        //SimpleDateFormat sdFormat = new SimpleDateFormat("yyyy/MM/dd");
        //Date current = new Date();

        String APPCASE  = Utils.convertToString(DBTools.dLookUp("APPCASE", "MBCASEAPP_M", "CASEID='" + wherestring[0] + "'", "SynctConn")); 
		String DIST_DESC = "";
		String STEP_DESC = "";


		if( !StringUtils.isEmpty(wherestring[1]) )
			 DIST_DESC  = Utils.convertToString(DBTools.dLookUp("APPCASE", "MBCASEAPP_M", "CASEID='" + wherestring[0] + "'", "SynctConn")); 

		if( !StringUtils.isEmpty(wherestring[2]) && wherestring[2].equals("1"))
			DIST_DESC = " ���ץ�";
		else if( !StringUtils.isEmpty(wherestring[2]) && wherestring[2].equals("4"))
			DIST_DESC = " �@��ץ�";
        	 

		if( !StringUtils.isEmpty(wherestring[3])){
	    	if (wherestring[3].equals("1")){
	    		STEP_DESC = " ��" + wherestring[3] + "���q�G���^���ץ����";
	    	}else if (wherestring[3].equals("2")){
	    		STEP_DESC = " ��" + wherestring[3] + "���q�G���^���ץ����";
	    	}else if (wherestring[3].equals("3")){
	    		STEP_DESC = " ��" + wherestring[3] + "���q�G���^���ץ���ӡ]��2���q�β�3���q���^����ơ^";
	    	}else if (wherestring[3].equals("4")){
	    		STEP_DESC = " ��" + wherestring[3] + "���q�G���^���ץ����";
	    	}
		}

        
        pageRow = sheet.getRow(page * page_rows);
        pageCell = pageRow.getCell((short)0);
        setBig5CellValue( "�s�_���I�u���ؿv�u�{" + APPCASE + DIST_DESC + STEP_DESC,pageCell);

	} 

    //��g����
	private void printFoot(String printpage) throws Exception{
		   //�g�X����
           //HSSFRow pageRow = sheet.getRow((page * page_rows) + 13);
           //HSSFCell pageCell = pageRow.getCell((short)0);
		   //setBig5CellValue(printpage ,pageCell);

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
            //System.err.println("bm20113.java: before getDataValue.");
            data=getDataValue(wherestring);
            //System.err.println("bm20113.java: end getDataValue.");

            //�ƻs���Y�˦�
 	        //header_style = copyPageHeaderStyle(sheet, 0,0,dtl_start_row,dtl_cols); // (int start_row, int start_col, int rows ��detail�}�l���header���C��, int cols)

 	        //�ƻs���Y�x�s���
            //header_value = copyPageHeaderValue(sheet, 0,0,dtl_start_row,dtl_cols);

            //�ƻs���˦�
            body_style = copyPageBodyStyleBlock(sheet, 2,0,3,dtl_cols);       // (int row, int start col, int cols)

             //System.err.println("bi30101.java: end body_style.");
           //�ƻs���x�s���
            body_value = copyPageBodyValueBlock(sheet, 2,0,3,dtl_cols);
           // System.err.println("bi30101.java: end body_value.");

            //�ƻs�Ĥ@���������X���x�s��
            region = copyMergedRegion(sheet);
               
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
                printHeader(wherestring);
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
         //System.err.println("AP30000:execOut error is "+e.toString());
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
					for (int i=0;i < dtl_cols;i++){
	             	    cell1 = row.getCell((short)(i));
	               	    setBig5CellValue(data1[i],cell1);
					}

               }catch(Exception e)     {
                  System.err.println("bm20113:printPageBody error is " + e);
               }

    }


 }