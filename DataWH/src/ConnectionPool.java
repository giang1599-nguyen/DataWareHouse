
import java.sql.SQLException;

import java.sql.Connection;

public interface ConnectionPool {
	public Connection getConnectionSql(String url, String username, String pass) throws SQLException;

}
