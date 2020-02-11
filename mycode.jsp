
import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;



public class XlsTest {
	public static void main(String[] args) throws  FileNotFoundException, IOException 
	{
		ArrayList<String []> reportData = new ArrayList<String []>();
		String [] temp =new String[2];
		temp[0] = "Name";
		temp[1] = "Name";
		reportData.add(temp);
		temp =new String[2];
		temp[0] = "EamilId";
		temp[1] = "test988@yopmail.com";
		reportData.add(temp);
		temp =new String[2];
		temp[0] = "Mobile";
		temp[1] = "546477";
		reportData.add(temp);
		temp =new String[2];
		temp[0] = "Country";
		temp[1] = "GB";
		reportData.add(temp);
		reportData.add(new String[1]);
		reportData.add(new String[1]);
		temp =new String[5];
		temp[0] = "Rank the following attributes (on a scale of 0 to 10) when using Smart Money";
		temp[1] = "Security";
		temp[2] = "Convenience";
		temp[3] = "Rates";
		temp[4] = "Transparency";
		reportData.add(temp);
		temp =new String[5];
		temp[0] = "";
		temp[1] = "10";
		temp[2] = "10";
		temp[3] = "10";
		temp[4] = "10";
		reportData.add(temp);
		reportData.add(new String[1]);
		reportData.add(new String[1]);
		temp =new String[7];
		temp[0] = "Please rate the following about BFC Smart Money";
		temp[1] = "Navigation";
		temp[2] = "Functionality";
		temp[3] = "LookNFeel";
		temp[4] = "Speed";
		temp[5] = "Stability";
		temp[6] = "Content";
		reportData.add(temp);
		temp =new String[7];
		temp[0] = "";
		temp[1] = "Very good";
		temp[2] = "Very good";
		temp[3] = "Very good";
		temp[4] = "Very good";
		temp[5] = "Very good";
		temp[6] = "Very good";
		reportData.add(temp);
		reportData.add(new String[1]);
		reportData.add(new String[1]);
		temp =new String[2];
		temp[0] = "In relation to the previous question, please provide comments on what you would like to be improved about BFC Smart Money";
		temp[1] = "fhdfhgfhgfhgh";
		reportData.add(temp);
		reportData.add(new String[1]);
		reportData.add(new String[1]);
		temp =new String[2];
		temp[0] = "How user-friendly is BFC Smart Moneys interface";
		temp[1] = "fhgfhgfhgf";
		reportData.add(temp);
		reportData.add(new String[1]);
		reportData.add(new String[1]);
		temp =new String[2];
		temp[0] = "Overall, how well does BFC Smart Money meet your needs";
		temp[1] = "fhgfhgfhgf";
		reportData.add(temp);
		reportData.add(new String[1]);
		reportData.add(new String[1]);
		temp =new String[2];
		temp[0] = "How likely is it that you would recommend BFC Smart Money to a friend or colleague";
		temp[1] = "10";
		reportData.add(temp);
		
		HSSFWorkbook hwb1=new HSSFWorkbook();
		hwb1 = downloadFileExcel(reportData);
		OutputStream op = new FileOutputStream("D:/Remit Project/feedBack.xls");
		hwb1.write(op);
		op.flush();

		//getdata();
		 /*try (OutputStream fileOut = new FileOutputStream("D:/Remit Project/Javatpoint2.xls")) {  
	            Workbook wb = new HSSFWorkbook();  
	            Sheet sheet = wb.createSheet("Sheet");
	            sheet.showInPane((short)0, (short)0);
				int d=0, columnCount=0;
	            if(reportData != null && reportData.size() > 1)
				{
	            	for(int x =0; x < reportData.size(); x++)
					{ 
	            		String [] temp11	= (String [])reportData.get(x);
	            		
						HSSFRow rowhead3=   (HSSFRow) sheet.createRow(d++);
						columnCount = temp11.length;
						for (int i=0; i<columnCount; i++)
						{
							HSSFCell cell03 = rowhead3.createCell(i);
							if(d==1)
							{
								HSSFCellStyle cellStyle = (HSSFCellStyle) wb.createCellStyle();
								HSSFFont font = (HSSFFont) wb.createFont();
								cellStyle.setFont(font);
								cell03.setCellStyle(cellStyle);
								cell03.setCellStyle(cellStyle);
								
							}
								cell03.setCellValue(temp[i]);
						}
						
					}
	            	//sheet.addMergedRegion(new CellRangeAddress(1,2,1,2));
				}
	            
	            Row row = sheet.createRow(1);  
	            Cell cell = row.createCell(1);  
	            cell.setCellValue("Two cells have merged");  
	              //Merging cells by providing cell index  
	            sheet.addMergedRegion(new CellRangeAddress(1,1,1,2));  
	            wb.write(fileOut);  
	        }catch(Exception e) {  
	            System.out.println(e.getMessage());  
	        }  */
	}
	public static HSSFWorkbook downloadFileExcel(ArrayList data)
	{
		HSSFWorkbook hwbDwnld=new HSSFWorkbook();
	
		try
		{
			if (data!= null && data.size()>0)
			{
				hwbDwnld = writeToFileExcel(data);
			}
			else
			{
			}
		}
		catch (Throwable t)
		{
		}

		data = null;
		return hwbDwnld;
	}
	public static HSSFWorkbook writeToFileExcel(ArrayList data)
	{	
		HSSFWorkbook hwb=new HSSFWorkbook();

		try
		{
			HSSFCellStyle style1 = hwb.createCellStyle();
			HSSFFont font1 = hwb.createFont();

			//style3.setFillForegroundColor(HSSFColor.WHITE.index); //HSSFColor.WHITE
			
			HSSFSheet sheet =  hwb.createSheet("FeedBack report");
			//set 1st row 1st column selected
			sheet.showInPane((short)0, (short)0);
			int d=0, columnCount=0;
			if(data != null && data.size() > 1)
			{
				for(int x =0; x < data.size(); x++)
				{
					String [] temp	= (String [])data.get(x);
					HSSFRow rowhead3=   sheet.createRow(d++);

					columnCount = temp.length;
					for (int i=0; i<columnCount; i++)
					{
							HSSFCell cell03 = rowhead3.createCell(i);
							cell03.setCellValue(temp[i]);
					}

						//sheet.addMergedRegion(new CellRangeAddress(1,1,1,2));
					
				} //
				for(short col=0; col<columnCount; col++)	//for auto size columns
					sheet.autoSizeColumn(col);			
				sheet.setSelected(true);
			}
		}
		catch ( Exception ex )
		{ 
		} 
		return hwb;
	}
}
