
import java.sql.SQLException;

import java.sql.Connection;

public interface ConnectionPool {
	public Connection getConnectionSql() throws SQLException;

}
