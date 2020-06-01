import java.io.FileInputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToDb {
	public static void main(String[] args) {
		String fileName = "E:\\\\ILoveYou\\\\DataWH\\\\Data_Information_17130047.xlsx";
		Vector dataHolder = read(fileName);
		saveToDatabase(dataHolder);
	}

	public static Vector read(String fileName) {
		Vector cellVectorHolder = new Vector();
		try {
			FileInputStream myInput = new FileInputStream(fileName);
			// POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
			XSSFWorkbook myWorkBook = new XSSFWorkbook(myInput);
			XSSFSheet mySheet = myWorkBook.getSheetAt(0);
			Iterator rowIter = mySheet.rowIterator();
			while (rowIter.hasNext()) {
				XSSFRow myRow = (XSSFRow) rowIter.next();
				Iterator cellIter = myRow.cellIterator();
				// Vector cellStoreVector=new Vector();
				List list = new ArrayList();
				while (cellIter.hasNext()) {
					XSSFCell myCell = (XSSFCell) cellIter.next();
					list.add(myCell);
				}
				cellVectorHolder.addElement(list);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return cellVectorHolder;
	}

	private static void saveToDatabase(Vector dataHolder) {
		String ho = "";
		String ten = "";
		String ngaysinh = "";
		String gioitinh = "";
		String nghe = "";
		String dchi = "";
		String email = "";
		System.out.println(dataHolder);

		for (Iterator iterator = dataHolder.iterator(); iterator.hasNext();) {
			List list = (List) iterator.next();
			ho = list.get(0).toString();
			System.out.println("họ "+ ho);
			ten = list.get(1).toString();
			ngaysinh = list.get(2).toString();
			gioitinh = list.get(3).toString();
			nghe = list.get(4).toString();
			dchi = list.get(5).toString();
			email = list.get(6).toString();

			try {
				Class.forName("com.mysql.jdbc.Driver").newInstance();
				Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/users?useUnicode=true&characterEncoding=utf-8", "root", "");
				System.out.println("connection made...");
				PreparedStatement stmt = con.prepareStatement(
						"INSERT INTO users(firstname, lastname, dateofbirth,gender,profession,address,email) VALUES (?, ?, ?, ?, ?, ?, ?)");
				stmt.setString(1, ho);
				stmt.setString(2, ten);
				stmt.setString(3, ngaysinh);
				stmt.setString(4, gioitinh);
				stmt.setString(5, nghe);
				stmt.setString(6, dchi);
				stmt.setString(7, email);
				stmt.executeUpdate();

				System.out.println("Data is inserted");
				stmt.close();
				con.close();
			} catch (ClassNotFoundException e) {
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			} catch (InstantiationException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			}
		}

	}
}