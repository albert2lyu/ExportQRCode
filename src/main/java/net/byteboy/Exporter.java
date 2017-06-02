package net.byteboy;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.WriterException;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;

public class Exporter {
//	部门数量
	private static final int DEPTMOUNT = 1;

	private ArrayList<String> info = new ArrayList<>();
	
	public static void main(String[] args) {
		
	}
	
	public void export(){
//		导入文件
		File excelFile = new File("");
		FileInputStream fis;
		try {
			fis = new FileInputStream(excelFile);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet[] sheets = new XSSFSheet[DEPTMOUNT];
			for(int i = 0; i < DEPTMOUNT; i++){
				sheets[i] = workbook.getSheetAt(i);
			}
			for(int i = 0; i < DEPTMOUNT; i++){
				for(int j = 0; j < sheets[i].getPhysicalNumberOfRows(); j++){
					readMessage(sheets[i], j);
					createQRCode();
				}
			}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public void createQRCode(){
				int width = 300;
				int height = 300;
				String imageFormat = "png";
				String contents = "姓名:"+info.get(0)+"\n学号:"+info.get(1)+"\n:"+info.get(2)+"\n职位:"+info.get(3)+"\n部门:"+info.get(4);
				
				HashMap<EncodeHintType, Object> hints = new HashMap<>();
				hints.put(EncodeHintType.CHARACTER_SET, "UTF-8");
				hints.put(EncodeHintType.MARGIN, 0);
				hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.M);
				
				try {
					BitMatrix bitMatrix = new MultiFormatWriter().encode(contents, BarcodeFormat.QR_CODE, width, height, hints);
					String dept = info.get(4);
					String id = info.get(1);
					String name = info.get(0);
//					生成文件
					File QRCodeFile = new File("");
					if(!QRCodeFile.exists()){
						QRCodeFile.mkdirs();
					}
					Path path = QRCodeFile.toPath();
					MatrixToImageWriter.writeToPath(bitMatrix, imageFormat, path);
				} catch (WriterException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
	}
	
	public void readMessage(XSSFSheet sheet, int i){
		Row row = sheet.getRow(i);
		int columns = row.getPhysicalNumberOfCells();
		Cell cell = null;
		for(int index = 0; index < columns; index++){
			cell = row.getCell(index);
			info.add(index, cell.getStringCellValue());
		}
	}
	
}
