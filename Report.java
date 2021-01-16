package com.report;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.Serializable;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Hashtable;

import javax.faces.bean.ManagedBean;
import javax.faces.bean.RequestScoped;
import javax.naming.Context;
import javax.naming.InitialContext;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.primefaces.model.DefaultStreamedContent;
import org.primefaces.model.StreamedContent;

import com.pardis.common.DateFormatUtils;
import com.report.constants.Constants;

@ManagedBean
@RequestScoped
public class Report  implements Serializable {
	private static final long serialVersionUID = 1L;

	
	 public StreamedContent getFile() throws IOException  {
		 
		 
		 	
		 	Connection con = null;  
			Statement stmt = null;  
			ResultSet rs = null;  
			

			Context ctx = null;
			Hashtable ht = new Hashtable();
			ht.put(Context.INITIAL_CONTEXT_FACTORY,  "weblogic.jndi.WLInitialContextFactory");
			ht.put(Context.PROVIDER_URL,Constants.SQL_CONNECTION.WL_DOMAIN);
			
			
			
			try {  
				ctx = new InitialContext(ht);
				javax.sql.DataSource ds = (javax.sql.DataSource) ctx.lookup(Constants.SQL_CONNECTION.DATASOURCENAME);
				con = ds.getConnection();
				stmt = con.createStatement();
				
				String query= "select *  from T_GeneralReports";
				rs = stmt.executeQuery(query);
				
				
				
				
				XSSFWorkbook wb_template = new XSSFWorkbook();
			    SXSSFWorkbook wb = new SXSSFWorkbook(wb_template); 
			    wb.setCompressTempFiles(true);
			    SXSSFSheet sh = (SXSSFSheet) wb.createSheet();
				
		    	//cache Size of Apachi Poi,
			    sh.setRandomAccessWindowSize(300000);
			    //Setting Header Values
				Row row = sh.createRow(0);
				row.createCell(1).setCellValue("سال");
				row.createCell(2).setCellValue("ماه");
				row.createCell(3).setCellValue("نام فروشگاه");
				row.createCell(4).setCellValue("نام خریدار");
				row.createCell(5).setCellValue("تاریخ خرید");
				row.createCell(6).setCellValue("خالص دریافتی");
				row.createCell(7).setCellValue("اقساط");
				row.createCell(8).setCellValue("حق بیمه");
				row.createCell(9).setCellValue("عوارض");
				row.createCell(10).setCellValue("حق بازاریابی");
				row.createCell(11).setCellValue("مالیات");
				
				Cell cell ;
				int rowNum = 1 ;
				
				//********Styling***************//
				CellStyle  styleCurrency=wb.createCellStyle();
				DataFormat format=wb.createDataFormat();
				styleCurrency.setDataFormat(format.getFormat("#,###"));
				
				Font font = wb.createFont();
				font.setBoldweight(Font.BOLDWEIGHT_BOLD);
				
				
				CellStyle  styleSummeryRow=wb.createCellStyle();
				styleSummeryRow.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
				styleSummeryRow.setFillPattern(CellStyle.SOLID_FOREGROUND);
				styleSummeryRow.setFont(font);
				
				
				CellStyle  styleYearSummery=wb.createCellStyle();
				styleYearSummery.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
				styleYearSummery.setFillPattern(CellStyle.SOLID_FOREGROUND);
				styleYearSummery.setFont(font);
				styleYearSummery.setDataFormat(format.getFormat("#,###"));
				
				
				CellStyle  styleMonthSummery=wb.createCellStyle();
				styleMonthSummery.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
				styleMonthSummery.setFillPattern(CellStyle.SOLID_FOREGROUND);
				styleMonthSummery.setFont(font);
				styleMonthSummery.setDataFormat(format.getFormat("#,###"));
				
				
				//Year of Current Loop
				String currentRowYear_Jalali="";
				//Year of Previous Loop 
				String previousRowYear_Jalali= ""; 
				//List Of Years which appears in result of select
				ArrayList< Integer> yearG= new ArrayList<Integer>();

				//***************************/
				
				//Merchant of Current Loop
				String currentRowMerchant="";
				//Merchant of Previous Loop 
				String previousRowMerchant= ""; 
				//List Of merchant which appears in result of select
				ArrayList< Integer> merchantG= new ArrayList<Integer>();
				ArrayList< String> merchanName= new ArrayList<String>();
				
				//***************************/
				//Month of Current Loop
				String currentRowMonth_PersianTxt="";
				//Month of Previous Loop 
				String previousRowMonth_PersianTxt= ""; 
				//List Of Months which appears in result of select
				ArrayList< Integer> monthG= new ArrayList<Integer>();
				
				//Month And Year In one Column
				String monthandYear ;
				Date buyDateGorgian;
				while (rs.next()) { 
					 
					buyDateGorgian = rs.getDate("crp_buydate");
					//Convert Georgina date to Jalali 
					String df= DateFormatUtils.getJalaliDate(buyDateGorgian);
					currentRowYear_Jalali=  df.substring(0, 4);
					currentRowMonth_PersianTxt=  df.substring(5, 7);
					 
					 
					 currentRowMerchant= rs.getString("crp_merchantcode");
					 monthandYear=  monthText( df);
					 row = sh.createRow(rowNum);
					
					 if (!previousRowYear_Jalali.equals(currentRowYear_Jalali )){
						 yearG.add(rowNum);
						 monthG.add(rowNum);
						 cell = row.createCell(0);
						 cell = row.createCell(1);
						 cell.setCellStyle( styleYearSummery);
						 cell.setCellValue( currentRowYear_Jalali);
						 previousRowYear_Jalali= currentRowYear_Jalali;
						 rowNum++;
						 row = sh.createRow(rowNum);
					
					 }
					 
					 if (!previousRowMerchant.equals(currentRowMerchant )){
						 merchantG.add(rowNum);
						 previousRowMerchant= currentRowMerchant;
						 
						 merchanName.add(rs.getString("crp_mrcname"));
					 }
					 
					 
					 if (!previousRowMonth_PersianTxt.equals(currentRowMonth_PersianTxt )){
						 
						 monthG.add(rowNum);
						 cell = row.createCell(2);
						 cell.setCellStyle( styleMonthSummery);
						 cell.setCellValue( currentRowMonth_PersianTxt);
						 previousRowMonth_PersianTxt= currentRowMonth_PersianTxt;
						
						
						 
						 rowNum++;
						 row = sh.createRow(rowNum);
						
						 
						 
					 }
					
					 
					row = sh.createRow(rowNum);
					cell = row.createCell(1);
					cell = row.createCell(2);
					cell.setCellValue(monthandYear);
					cell = row.createCell(3);
					cell.setCellValue( rs.getString("crp_mrcname"));
					cell = row.createCell(4);
					cell.setCellValue( rs.getString("crp_firstName") + " " +rs.getString("crp_lastName"));
					
					cell = row.createCell(5);
					cell.setCellValue( df);
					
					cell = row.createCell(6);
					cell.setCellValue( rs.getDouble("crp_prcpayedamount") );
					cell.setCellStyle(styleCurrency);
					
					cell = row.createCell(7);
					cell.setCellValue( rs.getDouble("crp_amount"));
					cell.setCellStyle(styleCurrency);
					
					cell = row.createCell(8);
					cell.setCellValue( rs.getDouble("crp_insur"));
					cell.setCellStyle(styleCurrency);
					
					cell = row.createCell(9);
					cell.setCellValue( rs.getDouble("crp_charge"));
					cell.setCellStyle(styleCurrency);
					
					cell = row.createCell(10);
					cell.setCellValue( rs.getDouble("crp_market"));
					cell.setCellStyle(styleCurrency);
					
					cell = row.createCell(11);
					cell.setCellValue( rs.getDouble("crp_tax"));
					cell.setCellStyle(styleCurrency);
					
					
					  rowNum++; 
				    
				}
				
				  for (int i= 0 ; i <10; i++){
				    	sh.setColumnWidth(i, 20*256);	
				    }
				
				 
				 yearG.add(rowNum+1);
				 monthG.add(rowNum);
				 merchantG.add(rowNum+1);
			
				
				for (int i =0 ; i< yearG.size()-1; i++){
					  
					  double sum1=0,sum2=0,sum3=0,sum4=0,sum5=0  ,sum6=0;
					  for (int j = yearG.get(i)+1; j<yearG.get(i+1);j++){
						if (monthG.contains(j)) continue; 
						double s1,s2,s3,s4,s5,s6; 
						try {
							 s1= sh.getRow(j).getCell(6).getNumericCellValue();
							 s2= sh.getRow(j).getCell(7).getNumericCellValue();
							 s3= sh.getRow(j).getCell(8).getNumericCellValue();
							 s4= sh.getRow(j).getCell(9).getNumericCellValue();
							 s5= sh.getRow(j).getCell(10).getNumericCellValue();
							 s6= sh.getRow(j).getCell(11).getNumericCellValue();
						} catch (Exception e) {
							 s1= 0 ; s2= 0 ; s3= 0 ; s3= 0 ; s4= 0 ; s5= 0 ; s6= 0 ;
						}
						sum1 = sum1 + s1;
						sum2 = sum2 + s2;
						sum3 = sum3 + s3;
						sum4 = sum4 + s4;
						sum5 = sum5 + s5;
						sum6 = sum6 + s6;
						}
					
					  row= sh.getRow(yearG.get(i)) ;
					  cell = row.createCell(2);  cell.setCellStyle(styleYearSummery);
					  cell = row.createCell(3);  cell.setCellStyle(styleYearSummery);
					  cell = row.createCell(4);  cell.setCellStyle(styleYearSummery);
					  cell = row.createCell(5);  cell.setCellStyle(styleYearSummery);
					  
					  cell = row.createCell(6); cell.setCellValue(sum1); cell.setCellStyle(styleYearSummery);
					  cell = row.createCell(7); cell.setCellValue(sum2); cell.setCellStyle(styleYearSummery);
					  cell = row.createCell(8); cell.setCellValue(sum3); cell.setCellStyle(styleYearSummery);
					  cell = row.createCell(9); cell.setCellValue(sum4); cell.setCellStyle(styleYearSummery);
					  cell = row.createCell(10); cell.setCellValue(sum5); cell.setCellStyle(styleYearSummery);
					  cell = row.createCell(11); cell.setCellValue(sum6); cell.setCellStyle(styleYearSummery);
					  
					
				  }
				  
				
				for (int i =0 ; i< monthG.size()-1; i++){
					if (yearG.contains(monthG.get(i))) continue; 
					
					
					  double sum1=0,sum2=0,sum3=0,sum4=0,sum5=0  ,sum6=0;
					  for (int j = monthG.get(i)+1; j<monthG.get(i+1);j++){
						 if (yearG.contains(j))continue; 
						 
						double s1,s2,s3,s4,s5,s6; 
						try {
							 s1= sh.getRow(j).getCell(6).getNumericCellValue();
							 s2= sh.getRow(j).getCell(7).getNumericCellValue();
							 s3= sh.getRow(j).getCell(8).getNumericCellValue();
							 s4= sh.getRow(j).getCell(9).getNumericCellValue();
							 s5= sh.getRow(j).getCell(10).getNumericCellValue();
							 s6= sh.getRow(j).getCell(11).getNumericCellValue();
						} catch (Exception e) {
							 s1= 0 ; s2= 0 ; s3= 0 ; s3= 0 ; s4= 0 ; s5= 0 ; s6= 0 ;
						}
						sum1 = sum1 + s1;
						sum2 = sum2 + s2;
						sum3 = sum3 + s3;
						sum4 = sum4 + s4;
						sum5 = sum5 + s5;
						sum6 = sum6 + s6;
						}
					  row= sh.getRow(monthG.get(i)) ;
					  cell = row.createCell(3);  cell.setCellStyle(styleMonthSummery);
					  cell = row.createCell(4);  cell.setCellStyle(styleMonthSummery);
					  cell = row.createCell(5);  cell.setCellStyle(styleMonthSummery);
					  
					  cell = row.createCell(6); cell.setCellValue(sum1); cell.setCellStyle(styleMonthSummery);
					  cell = row.createCell(7); cell.setCellValue(sum2); cell.setCellStyle(styleMonthSummery);
					  cell = row.createCell(8); cell.setCellValue(sum3); cell.setCellStyle(styleMonthSummery);
					  cell = row.createCell(9); cell.setCellValue(sum4); cell.setCellStyle(styleMonthSummery);
					  cell = row.createCell(10); cell.setCellValue(sum5); cell.setCellStyle(styleMonthSummery);
					  cell = row.createCell(11); cell.setCellValue(sum6); cell.setCellStyle(styleMonthSummery);
					  
				  }
		
			 
				for (int i=0 ; i< yearG.size()-1;i ++){
				    	sh.groupRow((yearG.get(i)+1), (yearG.get(i+1)-1));
				    	sh.setRowGroupCollapsed((yearG.get(i)+1), true);
				 }
				
			
				  
				 for (int i=0 ; i< monthG.size()-1;i ++){
				    	sh.groupRow((monthG.get(i)+1), (monthG.get(i+1)-1));
				    	sh.setRowGroupCollapsed(monthG.get(i)+1, true);
				 }
				  
				  sh.setRowSumsBelow(false);
			
				  
				 CellStyle  styleMerchantNameRow=wb.createCellStyle();
				 styleMerchantNameRow.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
					
				  for (int i=0 ; i< merchantG.size()-1;i ++){
				    	CellRangeAddress cellRangeAddress = new CellRangeAddress(merchantG.get(i)-1,merchantG.get(i+1)-2,0,0);
				    	sh.addMergedRegion(cellRangeAddress );
				    	cell= sh.getRow(merchantG.get(i)-1).createCell(0);
				    	cell.setCellValue(merchanName.get(i));
				    	cell.setCellStyle(styleMerchantNameRow);
				    	RegionUtil.setBorderTop(CellStyle.BORDER_MEDIUM, cellRangeAddress, sh, wb);
				    	RegionUtil.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex(), cellRangeAddress, sh, wb);
				 }
				
				  
				  
		
				 ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
				  wb.write(outputStream);

				  InputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());
				  StreamedContent   file =  new DefaultStreamedContent(inputStream, "application/xlsx", "AghsatVaKosuratePazirande.xlsx");
			    
				
				return file;
				
				
				
			}
				catch (Exception e) {  
					e.printStackTrace(); 
					return null;
				}  
				finally {  
					
					if (stmt != null) try { stmt.close(); } catch(Exception e) {}  
					if (con != null) try { con.close(); } catch(Exception e) {}
					if (rs != null) try { rs.close(); } catch(Exception e) {} 
				}
			
			
	    
	 }
	 public String monthText(String monthNumber){
		 String month =monthNumber.substring(5,7);
		 String year =monthNumber.substring(2,4);
		 
		 int m= Integer.valueOf(month);
		 switch (m) {
		case 1:
			return  "فروردین " +year;
		case 2:
			return "اردیبهشت " +year;
		case 3:
			return " خرداد"+ year;
		case 4:
			return "تیر "+ year;
		case 5:
			return "مرداد " +year;
		case 6:
			return "شهریور " +year;
		case 7:
			return "مهر " +year;
		case 8:
			return "آبان " + year;
		case 9:
			return "آذر " +year;
		case 10:
			return "دی " +year;
		case 11:
			return "بهمن " +year;
		case 12:
			return "اسفند " +year;
			

		default:
			return monthNumber;
		}
		
		 
	 }




}
