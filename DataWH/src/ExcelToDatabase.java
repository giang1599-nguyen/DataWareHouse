import java.io.*;
import java.sql.*;
import java.util.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelToDatabase {
	BufferedReader read;
	String driver = "jdbc:mysql:";
	String addressConfig = "localhost:3306/mydb";
	String addressSource = "localhost:3306/qlsv";
	String user = "root";
	String pass = "0126";
	Connection con;
	Connection conn;
	String tableName;
	String fileType;
	int numOfcol;
	int numOfdata;
	ArrayList<String> columnsList = new ArrayList<String>();
	int numberColumns;
	String dataPath;
	ArrayList<Cell> valuesList = new ArrayList<Cell>();

	public void createConnection() throws SQLException {
		System.out.println("Connecting database....");
		try {
			Class.forName("com.mysql.jdbc.Driver");
			con = DriverManager.getConnection(driver + "//" + addressConfig, user, pass);
			conn = DriverManager.getConnection(driver + "//" + addressSource, user, pass);

			System.out.println("Complete!!!!!");
			System.out.println("----------------------------------------------------------------");

		} catch (ClassNotFoundException e) {
			System.out.println("Can't connect!!!!!!!!!!");
			System.out.println("----------------------------------------------------------------");
		}

	}

	public void getConfig(int id, String delimiter) throws SQLException {
		PreparedStatement pre = (PreparedStatement) con
				.prepareStatement("SELECT * FROM mydb.configuration where id=?;");
		pre.setInt(1, id);
		ResultSet tmp = pre.executeQuery();
		tmp.next();
		tableName = tmp.getString("filename");
		fileType = tmp.getString("filetype");
		numOfcol = Integer.parseInt(tmp.getString("numofcol"));
		String listofcol = tmp.getString("listofcol");
		numOfdata = Integer.parseInt(tmp.getString("numofdata"));
		dataPath = tmp.getString("datapath");
		StringTokenizer tokens = new StringTokenizer(listofcol, delimiter);
		while (tokens.hasMoreTokens()) {
			columnsList.add(tokens.nextToken());
		}
		System.out.println("Get config: complete!!!!");

	}

	public void readData() throws IOException, EncryptedDocumentException, InvalidFormatException {
		File file = new File(dataPath);
		// URL url = new
		// URL("http://drive.ecepvn.org:5000/index.cgi?launchApp=SYNO.SDS.App.FileStation3.Instance&launchParam=openfile%3D%252FECEP%252Fsong.nguyen%252FDW_2020%252F&fbclid=IwAR1uHCCSN5m35GwuGLdNM0z41d6kn8IeXXaUBsv006I6EwbyZ8diTKLablg");
		FileInputStream fis = new FileInputStream(file);
		// InputStreamReader fis = new InputStreamReader(url.openStream());
		// creating workbook instance that refers to .xls file
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		// creating a Sheet object to retrieve the object
		XSSFSheet sheet = wb.getSheetAt(0);
		// evaluating cell type
		FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
		for (Row row : sheet) // iteration over row using for each loop
		{
			for (Cell cell : row) // iteration over cell using for each loop
			{
				valuesList.add(cell);

			}

		}
	}

	public String sqlCreatTable() {
		String preSql = "CREATE TABLE " + tableName + " (";
		preSql += columnsList.get(0) + " INT PRIMARY KEY NOT NULL,";
		for (int i = 1; i < numOfcol; i++) {
			preSql += columnsList.get(i) + " VARCHAR(100),";

		}
		preSql = preSql.substring(0, preSql.length() - 1) + ");";

		return preSql;

	}

	public void insertData(List<Cell> g) throws SQLException {

		String preSql = "INSERT INTO " + tableName + "(";
		for (int i = 0; i < numOfcol; i++) {
			preSql += columnsList.get(i) + ",";

		}
		preSql = preSql.substring(0, preSql.length() - 1) + ") VALUES ( ";

		int x = numOfcol;
		int y = x + numOfcol;
		PreparedStatement state;

		for (int i = 0; i < numOfdata; i++) {

			g = valuesList.subList(x, y);
			String a = preSql;
			a += g.get(0) + ",";
			for (int j = 1; j < g.size(); j++) {
				a += "\"" + g.get(j) + "\"" + ",";

			}
			a = a.substring(0, a.length() - 1) + ");";
			x += numOfcol;
			y = x + numOfcol;
			state = conn.prepareStatement(a);
			state.execute();
		}
	}

	public void extractData() throws SQLException {
		String sqlCreateTb = sqlCreatTable();
		// create table
		boolean tableStatus = false;
		boolean readDataStatus = false;
		// System.out.println(sqlCreateTb);

		try {
			System.out.println("Creating table " + tableName + ".......");
			PreparedStatement state = conn.prepareStatement(sqlCreateTb);
			state.execute();
			tableStatus = true;
			System.out.println("Create Table: Complete!!!");
			System.out.println("----------------------------------------------------------------");
		} catch (Exception e) {
			System.out.println("Can't create table " + tableName);
			System.out.println("----------------------------------------------------------------");
		}
		if (tableStatus) {
			try {
				readData();
				readDataStatus = true;
			} catch (Exception e) {
				System.out.println("Can't read data from " + dataPath);
			}
		}
		if (tableStatus && readDataStatus) {
			try {
				insertData(valuesList);
				System.out.println("Done!!!!!!!!!");
				System.out.println("----------------------------------------------------------------");

			} catch (Exception e) {
				System.out.println("Can't insert data!!!");
			}

		}

	}

	public static void main(String[] args)
			throws SQLException, IOException, EncryptedDocumentException, InvalidFormatException {
		ExcelToDatabase ex = new ExcelToDatabase();
		ex.createConnection();
		ex.getConfig(2, "|");
		ex.extractData();

		// ex.readData();
		// ex.insertData();
		// ex.convertDataToSql();
		// ex.sqlCreatTable();
		// ex.extractData();

		// ex.execute();
		// ex.ChangeWordToTxt();

		// ex.getTableInfor();
		// System.out.println(ex.numberColumns);
		// for (int i = 0; i < ex.ColumnList.size(); i++) {
		// System.out.print(ex.ColumnList.get(i)+ " ");
		// }
	}
}