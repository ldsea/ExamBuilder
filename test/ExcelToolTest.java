import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ildsea.bean.JudgeChoice;
import com.ildsea.bean.MutiChoice;
import com.ildsea.bean.SingleChoice;


public class ExcelToolTest {


	
	public void readExcelTest(String excelPath){
		File excelFile = null;// Excel文件对象  
        InputStream is = null;// 输入流对象  
        String cellStr = null;// 单元格，最终按字符串处理  
        try{
        	excelFile = new File(excelPath);  
            is = new FileInputStream(excelFile);// 获取文件输入流  
            XSSFWorkbook workbook2007 = new XSSFWorkbook(is);// 创建Excel2003文件对象  
            XSSFSheet sheet = workbook2007.getSheetAt(0);// 取出第一个工作表，索引是0   
            int rowNumber = sheet.getLastRowNum();//最大行

            
        }
        catch(IOException e){
        	e.printStackTrace();
        }
	}
	
	public List<SingleChoice> readExcel4Single(String excelPath){
		List<SingleChoice> list = new ArrayList<SingleChoice>();
		
		File excelFile = null;// Excel文件对象  
        InputStream is = null;// 输入流对象  
        String cellStr = null;// 单元格，最终按字符串处理  
        SingleChoice choice = null;
        try{
        	excelFile = new File(excelPath);  
            is = new FileInputStream(excelFile);// 获取文件输入流  
            XSSFWorkbook workbook2007 = new XSSFWorkbook(is);// 创建Excel2003文件对象  
            XSSFSheet sheet = workbook2007.getSheetAt(0);// 取出第一个工作表，索引是0   
            int rowNumber = sheet.getLastRowNum();
            for(int i=1;i<=rowNumber;i++){
            	choice = new SingleChoice();
            	XSSFRow row = sheet.getRow(i);// 获取行对象  
                if (row == null) {// 如果为空，不处理  
                    continue;  
                }
                for (int j = 0; j < row.getLastCellNum(); j++) {
                	XSSFCell cell = row.getCell(j);// 获取单元格对象
                	if (cell == null) {// 单元格为空设置cellStr为空串  
                        cellStr = "";  
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {// 对布尔值的处理  
                        cellStr = String.valueOf(cell.getBooleanCellValue());  
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {// 对数字值的处理  
                        cellStr = (int)cell.getNumericCellValue()+"";
                    } else {// 其余按照字符串处理  
                        cellStr = cell.getStringCellValue();  
                    }  
                	 // 下面按照数据出现位置封装到bean中  
                    if (j == 0) {  
                    	choice.setId(cellStr);
                    } else if (j == 1) {  
                        choice.setQuestion(cellStr);
                    } else if (j == 2) {  
                        choice.setA(cellStr); 
                    } else if (j == 3) {  
                    	choice.setB(cellStr); 
                    } else if (j==4){  
                    	choice.setC(cellStr); 
                    }  
                    else if(j==5){
                    	choice.setD(cellStr); 
                    }
                    else if(j==6){
                    	choice.setAnswer(cellStr);
                    }
                }
                list.add(choice);
            }            
        }
        catch(IOException e){
        	e.printStackTrace();
        }		
        finally {// 关闭文件流  
            if (is != null) {  
                try {  
                    is.close();  
                } catch (IOException e) {  
                    e.printStackTrace();  
                }  
            }  
        } 
		
		return list;
	}
	
	
	public List<MutiChoice> readExcel4Muti(String excelPath){
		List<MutiChoice> list = new ArrayList<MutiChoice>();
		
		File excelFile = null;// Excel文件对象  
        InputStream is = null;// 输入流对象  
        String cellStr = null;// 单元格，最终按字符串处理  
        MutiChoice choice = null;
        try{
        	excelFile = new File(excelPath);  
            is = new FileInputStream(excelFile);// 获取文件输入流  
            XSSFWorkbook workbook2007 = new XSSFWorkbook(is);// 创建Excel2003文件对象  
            XSSFSheet sheet = workbook2007.getSheetAt(1);// 取出第一个工作表，索引是0   
            int rowNumber = sheet.getLastRowNum();
            for(int i=1;i<=rowNumber;i++){
            	choice = new MutiChoice();
            	XSSFRow row = sheet.getRow(i);// 获取行对象  
                if (row == null) {// 如果为空，不处理  
                    continue;  
                }
                for (int j = 0; j < row.getLastCellNum(); j++) {
                	XSSFCell cell = row.getCell(j);// 获取单元格对象
                	if (cell == null) {// 单元格为空设置cellStr为空串  
                        cellStr = "";  
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {// 对布尔值的处理  
                        cellStr = String.valueOf(cell.getBooleanCellValue());  
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {// 对数字值的处理  
                        cellStr = (int)cell.getNumericCellValue()+"";
                    } else {// 其余按照字符串处理  
                        cellStr = cell.getStringCellValue();  
                    }  
                	 // 下面按照数据出现位置封装到bean中  
                    if (j == 0) {  
                    	choice.setId(cellStr);
                    } else if (j == 1) {  
                        choice.setQuestion(cellStr);
                    } else if (j == 2) {  
                        choice.setA(cellStr); 
                    } else if (j == 3) {  
                    	choice.setB(cellStr); 
                    } else if (j==4){  
                    	choice.setC(cellStr); 
                    }  
                    else if(j==5){
                    	choice.setD(cellStr); 
                    }
                    else if(j==6){
                    	choice.setAnswer(cellStr);
                    }
                }
                list.add(choice);
            }            
        }
        catch(IOException e){
        	e.printStackTrace();
        }		
        finally {// 关闭文件流  
            if (is != null) {  
                try {  
                    is.close();  
                } catch (IOException e) {  
                    e.printStackTrace();  
                }  
            }  
        } 
		
		return list;
	}	
	
	
	public List<JudgeChoice> readExcel4Judge(String excelPath){
		List<JudgeChoice> list = new ArrayList<JudgeChoice>();
		
		File excelFile = null;// Excel文件对象  
        InputStream is = null;// 输入流对象  
        String cellStr = null;// 单元格，最终按字符串处理  
        JudgeChoice choice = null;
        try{
        	excelFile = new File(excelPath);  
            is = new FileInputStream(excelFile);// 获取文件输入流  
            XSSFWorkbook workbook2007 = new XSSFWorkbook(is);// 创建Excel2003文件对象  
            XSSFSheet sheet = workbook2007.getSheetAt(2);// 取出第一个工作表，索引是0   
            int rowNumber = sheet.getLastRowNum();
            for(int i=1;i<=rowNumber;i++){
            	choice = new JudgeChoice();
            	XSSFRow row = sheet.getRow(i);// 获取行对象  
                if (row == null) {// 如果为空，不处理  
                    continue;  
                }
                for (int j = 0; j < row.getLastCellNum(); j++) {
                	XSSFCell cell = row.getCell(j);// 获取单元格对象
                	if (cell == null) {// 单元格为空设置cellStr为空串  
                        cellStr = "";  
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {// 对布尔值的处理  
                        cellStr = String.valueOf(cell.getBooleanCellValue());  
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {// 对数字值的处理  
                        cellStr = (int)cell.getNumericCellValue()+"";
                    } else {// 其余按照字符串处理  
                        cellStr = cell.getStringCellValue();  
                    }  
                	 // 下面按照数据出现位置封装到bean中  
                    if (j == 0) {  
                    	choice.setId(cellStr);
                    } else if (j == 1) {  
                        choice.setQuestion(cellStr);
                    } else if (j == 2) {  
                        choice.setAnswer(cellStr); 
                    }
                }
                list.add(choice);
            }            
        }
        catch(IOException e){
        	e.printStackTrace();
        }		
        finally {// 关闭文件流  
            if (is != null) {  
                try {  
                    is.close();  
                } catch (IOException e) {  
                    e.printStackTrace();  
                }  
            }  
        } 
		
		return list;
	}	
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		ExcelToolTest tool = new ExcelToolTest();
		List<JudgeChoice> list = tool.readExcel4Judge("/Users/ldsea/Desktop/exam.xlsx");
		for(int i=0;i<list.size();i++){
			System.out.println(list.get(i).toString()+"\n");
		}
	}



}
