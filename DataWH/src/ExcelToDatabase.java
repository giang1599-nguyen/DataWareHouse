import java.io.*;
import java.sql.*;
import java.util.*;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelToDatabase {
	GetConnectDatabase connect = new GetConnectDatabase();

	public ExcelToDatabase() throws SQLException {
		Connection connection = connect.getConnectionSql();
		String excelFilePath = "E:\\ILoveYou\\DataWH\\Data_Information_17130047.xlsx";
		int batchSize = 30;

		try {
			long start = System.currentTimeMillis();

			FileInputStream inputStream = new FileInputStream(excelFilePath);

			Workbook workbook = new XSSFWorkbook(inputStream);

			Sheet firstSheet = (Sheet) workbook.getSheetAt(0);
			Iterator<Row> rowIterator = firstSheet.iterator();

			connection.setAutoCommit(false);

			String sql = "INSERT INTO users (firstname, lastname, dateofbirth,gender,profession,address,email) VALUES (?, ?, ?, ?, ?, ?, ?)";
			PreparedStatement statement = connection.prepareStatement(sql);

			int count = 0;

			rowIterator.next(); // skip the header row

			while (rowIterator.hasNext()) {
				Row nextRow = rowIterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				while (cellIterator.hasNext()) {
					Cell nextCell = cellIterator.next();
					int columnIndex = nextCell.getColumnIndex() + 1;
					System.out.println(columnIndex);
					// try {
					switch (columnIndex) {

					// case 0:
					// String id = nextCell.getStringCellValue();
					// System.out.println(id);
					// statement.setString(1, id);
					// break;
					case 1:
						String fname = nextCell.getStringCellValue();
						statement.setString(1, fname);
						System.out.println(fname);
						break;
					case 2:
						String lname = nextCell.getStringCellValue();
						System.out.println(nextCell.getStringCellValue());
						statement.setString(2, lname);
						break;
					case 3:
						String enrollDate = nextCell.getStringCellValue();
						statement.setString(3, enrollDate);
						System.out.println(nextCell.getStringCellValue());
						break;
					case 4:
						String gender = nextCell.getStringCellValue();
						statement.setString(4, gender);
						break;
					case 5:
						String profes = nextCell.getStringCellValue();
						System.out.println(nextCell.getStringCellValue());
						statement.setString(5, profes);
						break;
					case 6:
						String addr = nextCell.getStringCellValue();
						System.out.println(nextCell.getStringCellValue());
						statement.setString(6, addr);
						break;
					case 7:
						String email = nextCell.getStringCellValue();
						System.out.println(nextCell.getStringCellValue());
						statement.setString(7, email);
						break;
					}
				}
				statement.addBatch();

				if (count % batchSize == 0) {
					statement.executeBatch();
				}

			}

			workbook.close();

			// execute the remaining queries
			statement.executeBatch();

			connection.commit();
			connection.close();

			long end = System.currentTimeMillis();
			System.out.printf("Import done in %d ms\n", (end - start));

		} catch (IOException ex1) {
			System.out.println("Error reading file");
			ex1.printStackTrace();
		} catch (SQLException ex2) {
			System.out.println("Database error");
			ex2.printStackTrace();
		}

	}

	public static void main(String[] args) throws SQLException {
		new ExcelToDatabase();
	}
}