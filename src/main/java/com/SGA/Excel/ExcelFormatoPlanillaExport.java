package com.SGA.Excel;

import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.SGA.entidades.Estudiante;

public class ExcelFormatoPlanillaExport {
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	
	private List<Estudiante> listStudent;
	
	
	public ExcelFormatoPlanillaExport(List<Estudiante> listStudent) {
		this.listStudent=listStudent;
		workbook = new XSSFWorkbook();
		
	}
	
	private void createCell(Row row,int columnCount, Object value,CellStyle style) {
		sheet.autoSizeColumn(columnCount);
		Cell cell=row.createCell(columnCount);
		if(value instanceof Long) {
			cell.setCellValue((Long) value);
		}else if(value instanceof Integer) {
			cell.setCellValue((Integer) value);
		}else if(value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		}else {
			cell.setCellValue((String) value);
		}
		cell.setCellStyle(style);
	}
	private void writeHeaderLine() {
		sheet=workbook.createSheet("Student");
		
		Row row = sheet.createRow(0);
		CellStyle style = workbook.createCellStyle();
		CellStyle style2 = workbook.createCellStyle();
		CellStyle style3 = workbook.createCellStyle();
		CellStyle style4 = workbook.createCellStyle();
		CellStyle style5 = workbook.createCellStyle();
		CellStyle style6 = workbook.createCellStyle();
		CellStyle style7 = workbook.createCellStyle();
		CellStyle style8 = workbook.createCellStyle();
		CellStyle style9 = workbook.createCellStyle();
		CellStyle style10 = workbook.createCellStyle();
		XSSFFont font=workbook.createFont();
		XSSFFont fuente=workbook.createFont();
		XSSFFont fuenteD=workbook.createFont();
		
		fuenteD.setFontHeight(12);
		fuenteD.setBold(true);
		style10.setAlignment(HorizontalAlignment.CENTER_SELECTION);
		style10.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	    style10.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style10.setBorderLeft(BorderStyle.MEDIUM);
		style10.setBorderRight(BorderStyle.MEDIUM);
		style10.setFont(fuenteD);
		
		
		
		
		
		style2.setAlignment(HorizontalAlignment.CENTER_SELECTION);
		
		style3.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		
	    style3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    style5.setBorderBottom(BorderStyle.MEDIUM);
	    style5.setBorderTop(BorderStyle.MEDIUM);
		font.setBold(true);
		fuente.setBold(true);
		font.setFontHeight(12);
		style7.setFont(fuente);
		style7.setBorderRight(BorderStyle.MEDIUM);
		style8.setFont(fuente);
		style8.setBorderLeft(BorderStyle.MEDIUM);
		style9.setFont(fuente);

		style9.setBorderBottom(BorderStyle.MEDIUM);
		style.setFont(font);
		style2.setFont(font);
		style3.setFont(font);
		style4.setFont(fuente);
		fuente.setFontHeight(10);
		style6.setBorderBottom(BorderStyle.MEDIUM);
		style6.setBorderRight(BorderStyle.MEDIUM);
		style6.setBorderTop(BorderStyle.MEDIUM);
		style3.setAlignment(HorizontalAlignment.CENTER_SELECTION);
		style3.setBorderRight(BorderStyle.MEDIUM);
		style4.setBorderRight(BorderStyle.MEDIUM);
		style4.setBorderLeft(BorderStyle.MEDIUM);
		style4.setRotation((short)90);
		style7.setRotation((short)90);
		style8.setRotation((short)90);
		style9.setRotation((short)90);
		style10.setRotation((short)90);
		style4.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	    style4.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    style7.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	    style7.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    style8.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	    style8.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		
		
		
		
		
		row=sheet.createRow(0);
        createCell(row,6,"REGISTRO Y CONTROL DE ASISTENCIA",style2);
        sheet.addMergedRegion(new CellRangeAddress(0,2,0,5));
        sheet.addMergedRegion(new CellRangeAddress(0,2,6,39)); 
        
        row=sheet.createRow(3);
        createCell(row,0,"DEPARTAMENTO/CODIGO DANE:HUILA/41",style2);
        createCell(row,3,"MUNICIPIO/CODIGO DANE:",style2);
        createCell(row,5,"INSTITUCION EDUCATIVA/CODIGO DANE:",style);
        createCell(row,20,"SEDE EDUCATIVA/CODIGO DANE:",style);
        createCell(row,36,"MES/AÑO:",style);
        sheet.addMergedRegion(new CellRangeAddress(3,4,0,2));
        sheet.addMergedRegion(new CellRangeAddress(3,4,3,4));
        sheet.addMergedRegion(new CellRangeAddress(3,4,5,19));
        sheet.addMergedRegion(new CellRangeAddress(3,4,20,35));
        sheet.addMergedRegion(new CellRangeAddress(3,4,36,39));
       
                     
        
        row=sheet.createRow(5);
        createCell(row,0,"CUPOS ALMUERZOS:",style2);
        createCell(row,3,"RACIONES PROGRAMADAS ALMUERZOS:",style);
        createCell(row,6,"CUPOS COMPLEMENTO ALIMENTARIO:",style);
        createCell(row,11,"RECIONES PROGRAMADAS C.A:AM______ PM______",style);
        createCell(row,27,"OPERADOR:",style);
        createCell(row,36,"No OPERACION:",style);
        sheet.addMergedRegion(new CellRangeAddress(5,6,3,5));
        sheet.addMergedRegion(new CellRangeAddress(5,6,11,26));
        sheet.addMergedRegion(new CellRangeAddress(5,8,27,35));
        sheet.addMergedRegion(new CellRangeAddress(5,8,36,39));
        sheet.addMergedRegion(new CellRangeAddress(5,8,0,2));
        sheet.addMergedRegion(new CellRangeAddress(5,8,6,10));
       
     
        sheet.addMergedRegion(new CellRangeAddress(7,8,3,5));
        sheet.addMergedRegion(new CellRangeAddress(7,8,11,26));
        sheet.addMergedRegion(new CellRangeAddress(9,9,0,26));
        sheet.addMergedRegion(new CellRangeAddress(10,11,17,38));
        sheet.addMergedRegion(new CellRangeAddress(10,15,39,39));
        sheet.addMergedRegion(new CellRangeAddress(10,10,10,12));
        sheet.addMergedRegion(new CellRangeAddress(11,11,10,12));

        

   
        sheet.addMergedRegion(new CellRangeAddress(11,15,13,13));
        sheet.addMergedRegion(new CellRangeAddress(11,15,16,16));

        sheet.addMergedRegion(new CellRangeAddress(11,15,9,9));
        sheet.addMergedRegion(new CellRangeAddress(12,15,12,12));
        sheet.addMergedRegion(new CellRangeAddress(12,15,11,11));
        sheet.addMergedRegion(new CellRangeAddress(12,15,10,10));
        sheet.addMergedRegion(new CellRangeAddress(11,15,14,14));
        sheet.addMergedRegion(new CellRangeAddress(11,15,15,15));
        sheet.addMergedRegion(new CellRangeAddress(57,61,0,14));
        sheet.addMergedRegion(new CellRangeAddress(62,67,0,2));
        sheet.addMergedRegion(new CellRangeAddress(62,67,6,7));
        sheet.addMergedRegion(new CellRangeAddress(62,63,3,5));
        sheet.addMergedRegion(new CellRangeAddress(64,65,3,5));
        sheet.addMergedRegion(new CellRangeAddress(66,67,3,5));
        
        sheet.addMergedRegion(new CellRangeAddress(62,63,8,21));
        sheet.addMergedRegion(new CellRangeAddress(64,65,8,21));
        sheet.addMergedRegion(new CellRangeAddress(66,67,8,21));

        
        sheet.addMergedRegion(new CellRangeAddress(62,67,22,29));
        
        sheet.addMergedRegion(new CellRangeAddress(62,63,30,38));
        sheet.addMergedRegion(new CellRangeAddress(64,65,30,38));
        sheet.addMergedRegion(new CellRangeAddress(66,67,30,38));

        
        row=sheet.createRow(7);
        createCell(row,3,"RACIONES ATENDIDAS ALMUERZOS:",style);
        createCell(row,11,"RACIONES ATENDIDAS C.A:AM______ PM______",style);
        
        
      
        
   
        
        
        row=sheet.createRow(10);
        int f=0;
		for (f = 0; f <=8 ; f++) {
			 createCell(row,f,"",style3);
		}
        createCell(row,9,"",style4);
        createCell(row,10,"",style3);
        createCell(row,11,"",style);
        createCell(row,12,"",style);
        createCell(row,13,"",style4);
        createCell(row,14,"",style4);
        createCell(row,15,"",style8);
        createCell(row,16,"",style7);
        createCell(row,17,"FECHA DE ENTREGA. Escriba el dia hábil al cual corresponde la entrega del complemento alimentairo",style3);
        createCell(row,39,"RACIONES ATENDIDAS",style10);
        
        
        row=sheet.createRow(11);
        createCell(row,0,"No",style3);
        createCell(row,1,"TIPO DOCUMENTO",style3);
        createCell(row,2,"No. DOCUMENTO DE",style3);
        createCell(row,3,"PRIMER APELLIDO TITULAR",style3);
        createCell(row,4,"SEGUNDO APELLIDO TITULAR",style3);
        createCell(row,5,"PRIMER NOMBRE TITULAR",style3);
        createCell(row,6,"SEGUNDO NOMBRE TITULAR",style3);
        createCell(row,7,"EDAD/FECHA DE",style3);
        createCell(row,8,"GENERO",style4);
        createCell(row,9,"GRADO EDUCATIVO",style4);
        createCell(row,10,"TIPO DE COMPLEMENTO",style);
        createCell(row,13,"PERTENENCIA ÉTNICA",style4);
        createCell(row,14,"VICTIMAS DEL CONFLICTO",style4);
        createCell(row,15,"EN CONDICION DE",style8);
        createCell(row,16,"DISCAPACIDAD",style7);
        createCell(row,39,"",style10);
      
        
       
        
        
        row=sheet.createRow(12);
        createCell(row,0,"",style3);
        createCell(row,1,"",style3);
        createCell(row,2,"IDENTIDAD",style3);
        createCell(row,3,"DE DERECHO",style3);
        createCell(row,4,"DE DERECHO",style3);
        createCell(row,5,"DE DERECHO",style3);
        createCell(row,6,"DE DERECHO",style3);
        createCell(row,7,"NACIMIENTO",style3);
        createCell(row,8,"",style3);
        createCell(row,9,"",style3);
        createCell(row,10,"ALMUERZO",style4);
        createCell(row,11,"AM",style4);
        createCell(row,12,"PM",style4);
        createCell(row,13,"",style4);
        createCell(row,14,"",style4);
        createCell(row,15,"",style8);
        createCell(row,16,"",style7);
        createCell(row,39,"",style10);
        
        
        
        row=sheet.createRow(13);
        int e=0;
		for (e = 0; e <=8 ; e++) {
			 createCell(row,e,"",style3);
		}
        createCell(row,9,"",style4);
        createCell(row,10,"",style4);
        createCell(row,11,"",style4);
        createCell(row,12,"",style4);
        createCell(row,13,"",style4);
        createCell(row,14,"",style4);
        createCell(row,15,"",style8);
        createCell(row,16,"",style7);
        createCell(row,39,"",style10);
       
        
        row=sheet.createRow(14);
        int d=0;
		for (d = 0; d <=8 ; d++) {
			 createCell(row,d,"",style3);
		}
    
        createCell(row,9,"",style4);
        createCell(row,10,"",style4);
        createCell(row,11,"",style4);
        createCell(row,12,"",style4);
        createCell(row,13,"",style4);
        createCell(row,14,"",style4);
        createCell(row,15,"",style8);
        createCell(row,16,"",style7);
        createCell(row,39,"",style10);

        createCell(row,17,"Días de atencion. Marque con una equis(x) el día que el titular de derecho recibe el complemento alimentario",style3);
        sheet.addMergedRegion(new CellRangeAddress(14,15,17,38));
        
        row=sheet.createRow(15);
        int c=0;
		for (c = 0; c <=8 ; c++) {
			 createCell(row,c,"",style3);
		}
     
        createCell(row,9,"",style4);
        createCell(row,10,"",style4);
        createCell(row,11,"",style4);
        createCell(row,12,"",style4);
        createCell(row,13,"",style4);
        createCell(row,14,"",style4);
        createCell(row,15,"",style8);
        createCell(row,16,"",style7);
        createCell(row,39,"",style10);

        
        row=sheet.createRow(57);
        createCell(row,0,"FIRMA DIARIA REPRESENTANTE SEDE EDUCATIVA",style);
        
        row=sheet.createRow(62);
        createCell(row,0,"RESPONSABLE SEDE EDUCATIVA:",style);
        createCell(row,3,"NOMBRE:",style);
        createCell(row,30,"NOMBRE:",style);
        createCell(row,6,"RECTOR INSTITUCION EDUCATIVA:",style);
        createCell(row,22,"RESPONSABLE DEL OPERADOR:",style);
       
        row=sheet.createRow(64);
        createCell(row,3,"FIRMA:",style);
        createCell(row,30,"FIRMA:",style);
        
        
        row=sheet.createRow(66);
        createCell(row,3,"CEDULA:",style);
        createCell(row,30,"CEDULA:",style);
        int i=16;
		for (i = 16; i <=61 ; i++) {
			sheet.addMergedRegion(new CellRangeAddress(i,i,15,16));
		}
		
		
		
		int a=17;
			for (a = 17; a <=38 ; a++) {
				sheet.addMergedRegion(new CellRangeAddress(12,13,a,a));
			}
			

	}
	private void writeDataLines() {
		int rowCount=16;
		
		CellStyle style=workbook.createCellStyle();
		XSSFFont font=workbook.createFont();
		font.setFontHeight(12);
		style.setFont(font);
		
		for(Estudiante stu:listStudent) {
			Row row=sheet.createRow(rowCount++);
			int columnCount=0;
		
			createCell(row, columnCount++, stu.getApellido1(), style);
			createCell(row, columnCount++, stu.getApellido2(), style);
			createCell(row, columnCount++, stu.getNombre1(), style);
			createCell(row, columnCount++, stu.getAcudiente(), style);
			createCell(row, columnCount++, stu.getDireccionRecidencia(), style);
		}
	}
	
	public void export(HttpServletResponse response) throws IOException{
		writeHeaderLine();
		writeDataLines();
	
		
		ServletOutputStream outputStream=response.getOutputStream();
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
	}
	
	
}
