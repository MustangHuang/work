package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Scanner;

import org.apache.commons.collections4.bag.SynchronizedSortedBag;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class First {
	public static final int intMax = Integer.MAX_VALUE;
	public static void main(String[] args) {
		//����double���� С��λ
		DecimalFormat df = new DecimalFormat("#.00");
		//���ڴ�� ���ԭʼֵ �� ��ֺ��ֵ
		Double[] dou = {0.00,0.00};
		//���ڸ��²�ֺ����һ�����ݵĽ����ֱ�Ӹ�ԭʼmap�ᵼ��ȫ�����ݸı�
		Map m = new LinkedHashMap<>();
		Workbook wb = null;
		Sheet sheet = null;
		Row row = null;
		List<Map<String,String>> list= null;
		String cellData = null;
		String filePath = "D:\\test.xls";
		String[] columns = {"num","count","money","flag"};
		wb = readExcel(filePath);
		int count = 1;
		if(wb != null) {
			//���ڴ�ű�������
			list = new ArrayList<Map<String,String>>();
			//��ȡ��һ��sheet
			sheet = wb.getSheetAt(0);
			//��ȡ�������
			int rownum = sheet.getPhysicalNumberOfRows();
			//��ȡ��һ��
			row = sheet.getRow(0);
			//��ȡ�������
			int colnum = row.getPhysicalNumberOfCells();
			for(int i = 0;i<rownum;i++) {
				Map<String,String> map = new LinkedHashMap<String,String>();
				row = sheet.getRow(i);
				if(row != null) {
					for(int j=0;j<colnum;j++) {
						cellData = (String)getCellFormatValue(row.getCell(j));
						if(columns[j].equals("num")) {
							//���� . ��Ҫת��
							cellData = cellData.split("\\.")[0];
						}else if(columns[j].equals("count")) {
							//double a = (double)cellData;		
							count = Integer.valueOf(cellData.substring(0,cellData.lastIndexOf(".")));
							cellData = "1";
						}else if(columns[j].equals("money")) {
							dou[0] = Double.valueOf(cellData);
							double result = Double.valueOf((String)cellData)/count;
							cellData = df.format(result);
							dou[1] = Double.valueOf(cellData)*(count - 1);
						}
						map.put(columns[j], cellData);
					}
				}else break;
				if(count == 1) {
					list.add(map);
				}else {
					for(int q = 0;q<count;q++) {
						if(q == (count -1)) {
							m = copyMap(map);
							m.replace("money", df.format(dou[0]-dou[1]));
							list.add(m);
							break;
						}
						list.add(map);
					}
				}
			}
		}
		exportExcel(list, filePath,columns);
		
	}
	
	//��ȡexcel
	public static Workbook readExcel(String filePath) {
		Workbook wb = null;
		if(filePath == null) {
			return null;
		}
		int len = filePath.lastIndexOf(".");
		String extString = filePath.substring(len);
		InputStream is = null;
		try {
			is = new FileInputStream(filePath);
			if(".xls".equals(extString)) {
				return wb = new HSSFWorkbook(is);
			}else if(".xlsx".equals(extString)) {
				return wb = new XSSFWorkbook(is);
			}else {
				return null;
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return wb;
	}
	
	public static Object getCellFormatValue(Cell cell) {
		Object cellValue = null;
		if(cell != null) {
			int i = cell.getCellType();
			//�ж�cell����
			if(Cell.CELL_TYPE_NUMERIC == i) {
				cellValue=String.valueOf(cell.getNumericCellValue());
//				cellValue=cell.getNumericCellValue();
			}else if(Cell.CELL_TYPE_FORMULA == i) {
				//�ж�cell�Ƿ�Ϊ���ڸ�ʽ
				if(DateUtil.isCellDateFormatted(cell)) {
					cellValue = cell.getDateCellValue();
				}else {
					//ת��Ϊ���ڸ�ʽYYYY-mm-dd
					cellValue = String.valueOf(cell.getNumericCellValue());
//					cellValue = cell.getNumericCellValue();
				}
			}else if(Cell.CELL_TYPE_STRING == i) {
				//����
				cellValue = cell.getRichStringCellValue().getString();
			}
		}else {
			cellValue = "";
		}
		return cellValue;
	}
	
	public static Map copyMap(Map<String,String> map) {
		Map<String,String> m = new LinkedHashMap<>();
		for (String set : map.keySet()) {
			m.put(set, map.get(set));
		}
		return m;
	}
	
	public static void exportExcel(List<Map<String,String>> list,String filePath,String[] columns) {
		Workbook wb = null;
		int len = filePath.lastIndexOf(".");
		String extString = filePath.substring(len);
		if(".xls".equals(extString)) {
			wb = new HSSFWorkbook();
		}else {
			wb = new XSSFWorkbook();
		}
		
		Sheet sheet = wb.createSheet();
		for(int i = 0;i<list.size();i++) {
			Row row = sheet.createRow(i);
			Map<String,String> map = list.get(i);
			for(int j = 0;j<columns.length;j++) {
				row.createCell(j).setCellValue(map.get(columns[j]));
			}
		}
		
		OutputStream ops = null;
		try {
			File f = new File("D:\\aaaaaaaaaaa"+extString);
			if(f.exists() && !f.isDirectory()) {
				System.out.println("�ļ��Ѵ��ڣ��Ƿ񸲸ǣ�y/n");
				Scanner sc = new Scanner(System.in);
				String yesOrNo = sc.next();
				if(yesOrNo.equals("y")) {
					ops = new FileOutputStream("D:\\aaaaaaaaaaa"+extString);
					wb.write(ops);
				}else {
					return;
				}
			}else {
				ops = new FileOutputStream("D:\\aaaaaaaaaaa"+extString);
				wb.write(ops);
			}
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally {
			try {
					if(ops != null) {
						ops.flush();
						ops.close();
					}
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			try {	
					if(wb != null) {
						wb.close();
					}
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		
	}
}
