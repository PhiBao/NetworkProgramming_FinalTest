package model;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import util.MySQLConnection;

public class GetResult {

	private Connection conn;
	private Statement statement1;
	private Statement statement2;

	public void export(String excelFilePath) {

		try {
			conn = MySQLConnection.connect();
			String sql1 = "SELECT PhongThi FROM phongthi";
			String sql2 = "SELECT MaGV FROM canbo";

			statement1 = conn.createStatement();
			statement2 = conn.createStatement();
			ResultSet result1 = statement1.executeQuery(sql1);
			ResultSet result2 = statement2.executeQuery(sql2);

			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Result");

			writeHeaderLine(sheet);

			writeDateLines(result1, result2, workbook, sheet);

			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);

			workbook.close();
			statement1.close();
			statement2.close();

			System.out.println("Result in " + excelFilePath);

		} catch (SQLException e) {
			System.out.println("Datababse error:");
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println("File IO error:");
			e.printStackTrace();
		}
	}

	private void writeHeaderLine(XSSFSheet sheet) {

		Row headerRow = sheet.createRow(0);

		Cell headerCell = headerRow.createCell(0);
		headerCell.setCellValue("STT");
		CellUtil.setAlignment(headerCell, HorizontalAlignment.CENTER);

		headerCell = headerRow.createCell(1);
		headerCell.setCellValue("Phòng thi");
		CellUtil.setAlignment(headerCell, HorizontalAlignment.CENTER);

		headerCell = headerRow.createCell(2);
		headerCell.setCellValue("Mã GV");
		CellUtil.setAlignment(headerCell, HorizontalAlignment.CENTER);

	}

	private void writeDateLines(ResultSet result1, ResultSet result2, XSSFWorkbook workbook, XSSFSheet sheet)
			throws SQLException {
		int rowCount = 1;

		while (result1.next()) {
			result2.next();
			String phongThi = result1.getString("PhongThi");
			String maGV = String.valueOf(result2.getBigDecimal("MaGV").longValue());

			Row row = sheet.createRow(rowCount++);
			int columnCount = 0;

			Cell cell = row.createCell(columnCount++);
			cell.setCellValue(rowCount - 1);
			CellUtil.setAlignment(cell, HorizontalAlignment.CENTER);

			cell = row.createCell(columnCount++);
			cell.setCellValue(phongThi);
			CellUtil.setAlignment(cell, HorizontalAlignment.CENTER);

			cell = row.createCell(columnCount++);
			cell.setCellValue(maGV);
			CellUtil.setAlignment(cell, HorizontalAlignment.CENTER);

		}
	}

}
