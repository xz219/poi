import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
	public static void main(String[] args) throws Exception {
		//2007
		String source = "C:\\Users\\Administrator\\Desktop\\test.xlsx";
//		String source = "C:\\Users\\Administrator\\Desktop\\ZkRyjbxx.xls";
		//����Excel�ļ�
		if(source.endsWith(".xlsx")){
			parseXlsx(source);
		}else if(source.endsWith(".xls")){
			parseXls(source);
			}else{
				throw new Exception("�ļ���ʽ����!!!");
			}
		}
	  // ����2003���excel�ļ�   
	private static void parseXls(String source)throws Exception{
		FileInputStream stream = new FileInputStream(source);
		HSSFWorkbook workbook = new HSSFWorkbook(stream);
//	     2�����ص�һ��sheetҳ
			HSSFSheet sheet = workbook.getSheetAt(0);
			for (Row row : sheet) {
				int rowNum = row.getRowNum();
				if(rowNum ==0){
					//������һ��
					continue;
				}
				String name = row.getCell(0).getStringCellValue();
				String age = row.getCell(1).getStringCellValue();
				String date = row.getCell(1).getStringCellValue();
				System.out.println("���� : " +name + "���� : "+age +" ���� : " +date);
	}
	}
//	����2007���excel�ļ�  
	private static void parseXlsx(String source)throws Exception{
		FileInputStream stream = new FileInputStream(source);
		 XSSFWorkbook xssfWorkbook = new XSSFWorkbook(stream); 
	      
		    // ѭ��������Sheet  
		    for(int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++){  
		      XSSFSheet xssfSheet = xssfWorkbook.getSheetAt( numSheet);  
		      if(xssfSheet == null){  
		        continue;  
		      }  
		        
		      // ѭ����Row   
		      for(int rowNum = 0; rowNum <= xssfSheet.getLastRowNum(); rowNum++ ){  
		        XSSFRow xssfRow = xssfSheet.getRow( rowNum);  
		        if(xssfRow == null){  
		          continue;  
		        }  
		          
		        // ѭ����Cell     
		        for(int cellNum = 0; cellNum <= xssfRow.getLastCellNum(); cellNum++){  
		          XSSFCell xssfCell = xssfRow.getCell( cellNum);  
		          if(xssfCell == null){  
		            continue;  
		          }  
		          System.out.println(xssfCell);  
		        }  
		      }  
		    }  
	}

}
