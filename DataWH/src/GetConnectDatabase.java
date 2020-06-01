import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class GetConnectDatabase implements ConnectionPool {

	public GetConnectDatabase() {

	}

	public Connection getConnectionSql() throws SQLException {

		String hostName = "localhost";
		String dbName = "users";
		String userName = "root";
		String password = "";
		String connectionURL = "jdbc:mysql://" + hostName + ":3306/" + dbName
				+ "?useUnicode=true&characterEncoding=utf-8";
		Connection conn = DriverManager.getConnection(connectionURL, userName, password);
		return conn;
	}

	public static void main(String[] args) {
		ConnectionPool cp = new GetConnectDatabase();
		try {
			Connection con = cp.getConnectionSql();
			if (con != null) {
				System.out.println(con);
			} else {
				System.out.println("connection is null");
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
}
