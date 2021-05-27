package com.test.poi.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

/**
 * Clase util para leer un excel y transformarlo a un hashmap.
 * 
 * 
 * @author Nelson Alvarado
 *
 */
public class ExcelUtil {

	private static final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

	private static final String XLS = ".xls";
	private static final String XLSX = ".xlsx";

	/**
	 * Obtenga el objeto de libro correspondiente según el sufijo del archivo
	 * 
	 * @param filePath
	 * @param fileType
	 * @return
	 */
	public static Workbook getWorkbook(String filePath, String fileType) {
		Workbook workbook = null;
		FileInputStream fileInputStream = null;
		try {
			File excelFile = new File(filePath);
			if (!excelFile.exists()) {
				logger.info(filePath + "El archivo no existe");
				return null;
			}
			fileInputStream = new FileInputStream(excelFile);
			if (fileType.equalsIgnoreCase(XLS)) {
				workbook = new HSSFWorkbook(fileInputStream);
			} else if (fileType.equalsIgnoreCase(XLSX)) {
				workbook = new XSSFWorkbook(fileInputStream);
			}
			
		} catch (Exception e) {
			logger.error("No se pudo obtener el archivo", e);
		} finally {
			try {
				if (null != fileInputStream) {
					fileInputStream.close();
				}
			} catch (Exception e) {
				logger.error("Error al cerrar el flujo de datos! Mensaje de error:", e);
				return null;
			}
		}
		return workbook;
	}
	
	
	/**
	 * Obtenga el objeto de libro correspondiente según el sufijo del archivo
	 * 
	 * @param filePath
	 * @param fileType
	 * @return
	 */
	public static Workbook getWorkbook(InputStream excelFile, String fileType) {
		Workbook workbook = null;
		try {
			if (fileType.equalsIgnoreCase(XLS)) {
				workbook = new HSSFWorkbook(excelFile);
			} else if (fileType.equalsIgnoreCase(XLSX)) {
				workbook = new XSSFWorkbook(excelFile);
			}
			
		} catch (Exception e) {
			logger.error("No se pudo obtener el archivo", e);
		} finally {
			try {
				if (null != excelFile) {
					excelFile.close();
				}
			} catch (Exception e) {
				logger.error("Error al cerrar el flujo de datos! Mensaje de error:", e);
				return null;
			}
		}
		return workbook;
	}
	
	
	
	/**
	 * Leer archivos de Excel en lotes y devolver objetos de datos
	 * 
	 * @param filePath
	 * @return
	 */
	public static List<Map<String, String>> readExcelMultipart(MultipartFile file) {
		Workbook  workbook = null;
		List<Map<String, String>> resultList = new ArrayList<>();
		String filePath = file.getOriginalFilename();
		
		try {
			String fileType = filePath.substring(filePath.lastIndexOf("."));
			 workbook = getWorkbook(file.getInputStream(), fileType);
			if (workbook == null) {
				logger.info("No se pudo obtener el objeto del libro de trabajo");
				return null;
			}
			resultList = analysisExcel(workbook);
			return resultList;
		} catch (Exception e) {
			logger.error("Error al leer el archivo de Excel" + filePath + "mensaje de error", e);
			return null;
		} finally {
			try {
				if (null != workbook) {
					//workbook.close();
				}
			} catch (Exception e) {
				logger.error("Error al cerrar el flujo de datos! Mensaje de error:", e);
				return null;
			}

		}
	}
	

	public static List<Object> readFolder(String filePath) {
		int fileNum = 0;
		File file = new File(filePath);
		List<Object> returnList = new ArrayList<>();
		List<Map<String, String>> resultList = new ArrayList<>();
		if (file.exists()) {
			File[] files = file.listFiles();
			for (File file2 : files) {
				if (file2.isFile()) {
					resultList = readExcel(file2.getAbsolutePath());
					returnList.add(resultList);
					fileNum++;
				}
			}
		} else {
			logger.info("La carpeta no existe");
			return null;
		}
		logger.info("Archivo común:" + fileNum);
		return returnList;
	}

	/**
	 * Leer archivos de Excel en lotes y devolver objetos de datos
	 * 
	 * @param filePath
	 * @return
	 */
	public static List<Map<String, String>> readExcel(String filePath) {
		Workbook  workbook = null;
		List<Map<String, String>> resultList = new ArrayList<>();
		try {
			String fileType = filePath.substring(filePath.lastIndexOf("."));
			 workbook = getWorkbook (filePath, fileType);
			if (workbook == null) {
				logger.info("No se pudo obtener el objeto del libro de trabajo");
				return null;
			}
			resultList = analysisExcel(workbook);
			return resultList;
		} catch (Exception e) {
			logger.error("Error al leer el archivo de Excel" + filePath + "mensaje de error", e);
			return null;
		} finally {
			try {
				if (null != workbook) {
					//workbook.close();
				}
			} catch (Exception e) {
				logger.error("Error al cerrar el flujo de datos! Mensaje de error:", e);
				return null;
			}

		}
	}

	/**
	 * Analizar el archivo de Excel y devolver el objeto de datos
	 * 
	 * @param workbook
	 * @return
	 */
	public static List<Map<String, String>> analysisExcel(Workbook workbook) {
		List<Map<String, String>> dataList = new ArrayList<>();
		int sheetCount = workbook.getNumberOfSheets(); // o tomar el número de hojas en Excel
		for (int i = 0; i < sheetCount; i++) {
			Sheet sheet = workbook.getSheetAt(i);

			if (sheet == null) {
				continue;
			}
			int firstRowCount = sheet.getFirstRowNum(); // Obtener el número de serie de la primera fila
			Row firstRow = sheet.getRow(firstRowCount);
			int cellCount = firstRow.getLastCellNum(); // Obtener el número de columnas

			List<String> mapKey = new ArrayList<>();

			// Obtenga la información del encabezado y colóquela en la Lista para su uso
			// posterior
			if (firstRow == null) {
				logger.info("No se pudo analizar Excel, no se leyeron datos en la primera fila");
			} else {
				for (int i1 = 0; i1 < cellCount; i1++) {
					mapKey.add(firstRow.getCell(i1).toString());
				}
			}

			// Analiza cada fila de datos para formar un objeto de datos
			int rowStart = firstRowCount + 1;
			int rowEnd = sheet.getPhysicalNumberOfRows();
			for (int j = rowStart; j < rowEnd; j++) {
				Row row = sheet.getRow(j); // Obtiene el objeto de fila correspondiente

				if (row == null) {
					continue;
				}

				Map<String, String> dataMap = new HashMap<>();
				// Convierta cada fila de datos en un objeto Map
				dataMap = convertRowToData(row, cellCount, mapKey);
				dataList.add(dataMap);
			}
		}
		return dataList;
	}

	/**
	 * Convierta cada fila de datos en un objeto Mapa
	 * 
	 * @param objeto    de fila de fila
	 * @param cellCount número de columnas
	 * @param mapKey    encabezado Mapa
	 * @return
	 */
	public static Map<String, String> convertRowToData(Row row, int cellCount, List<String> mapKey) {
		if (mapKey == null) {
			logger.info("Sin información de encabezado");
			return null;
		}
		Map<String, String> resultMap = new HashMap<>();
		Cell cell = null;
		for (int i = 0; i < cellCount; i++) {
			cell = row.getCell(i);
			if (cell == null) {
				resultMap.put(mapKey.get(i), "");
			} else {
				resultMap.put(mapKey.get(i), getCellVal(cell));
			}
		}
		return resultMap;
	}

	/**
	 * Obtener el valor de la celda
	 * 
	 * @param cel
	 * @return
	 */
	public static String getCellVal(Cell cel) {
		if (cel.getCellType() == Cell.CELL_TYPE_STRING) {
			return cel.getRichStringCellValue().getString();
		}
		if (cel.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			return cel.getNumericCellValue() + "";
		}
		if (cel.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
			return cel.getBooleanCellValue() + "";
		}
		if (cel.getCellType() == Cell.CELL_TYPE_FORMULA) {
			return cel.getCellFormula() + "";
		}
		return cel.toString();
	}

}
