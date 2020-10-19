package server;

import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import util.MySQLConnection;

public class OfficersPost {

	private Connection conn;
	private PreparedStatement ps;

	public void insertListOfficers(String excelFilePath) {
		try {

			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workBook = new XSSFWorkbook(inputStream);

			Sheet firstSheet = workBook.getSheetAt(0);
			Iterator<Row> rows = firstSheet.iterator();

			conn = MySQLConnection.connect();
			// Sét tự động commit false, để chủ động điều khiển
			conn.setAutoCommit(false);
			// skip the header row
			rows.next();

			String sql = "INSERT INTO canbo(MaGV, HoTen, NgaySinh, DonVi) VALUES (?, ?, ?, ?)";
			ps = conn.prepareStatement(sql);

			while (rows.hasNext()) {
				Row row = rows.next();
				Iterator<Cell> cells = row.cellIterator();
				boolean check = false;

				while (cells.hasNext()) {
					Cell cell = cells.next();
					check = isCellEmpty(cell);

					int columnIndex = cell.getColumnIndex();
					DataFormatter df = new DataFormatter();

					switch (columnIndex) {
					case 1:
						ps.setString(1, df.formatCellValue(cell));
						break;
					case 2:
						ps.setString(2, df.formatCellValue(cell));
						break;
					case 3:
						ps.setString(3, df.formatCellValue(cell));
						break;
					case 4:
						ps.setString(4, df.formatCellValue(cell));
						break;
					default:
						break;
					}
				}
				if (check == false)
					ps.addBatch();
				check = false;
			}

			workBook.close();
			// execute the remaining queries
			ps.executeBatch();
			ps.close();

			conn.commit();
			conn.rollback();

			System.out.println("Record is inserted into CANBO table!");

		} catch (Exception e) {
			e.printStackTrace();
			MySQLConnection.rollbackQuietly(conn);

		} finally {

			try {
				if (ps != null)
					ps.close();
			} catch (SQLException e) {
				e.printStackTrace();
			}

			MySQLConnection.disconnect(conn);
		}
	}

	private boolean isCellEmpty(final Cell cell) {
		if (cell == null) {
			return true;
		}

		if (cell.getCellType() == CellType.BLANK) {
			return true;
		}

		if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty()) {
			return true;
		}

		return false;
	}

}
