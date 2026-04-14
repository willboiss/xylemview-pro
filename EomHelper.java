import java.sql.*;

/**
 * Tiny helper for XylemView Pro — runs SQL against AS/400 via jt400 JDBC.
 * Usage: java -cp ".;E:\javaprod\jt400v7\lib\jt400.jar" EomHelper <mode> <sql>
 *   mode = "query" (SELECT → JSON rows) or "update" (UPDATE/INSERT → affected count)
 * Output: JSON to stdout. Errors go to stderr.
 */
public class EomHelper {
    private static final String HOST = "10.63.26.203";
    private static final String USER = "HTJAVA";
    private static final String PASS = "PACEQ21P";

    public static void main(String[] args) {
        if (args.length < 2) {
            System.err.println("Usage: EomHelper <query|update> <sql>");
            System.exit(1);
        }
        String mode = args[0];
        String sql = args[1];
        String url = "jdbc:as400://" + HOST + ";naming=system;errors=full";

        try {
            Class.forName("com.ibm.as400.access.AS400JDBCDriver");
            try (Connection conn = DriverManager.getConnection(url, USER, PASS)) {
                if ("update".equals(mode)) {
                    try (Statement st = conn.createStatement()) {
                        int affected = st.executeUpdate(sql);
                        System.out.println("{\"ok\":true,\"affected\":" + affected + "}");
                    }
                } else {
                    try (Statement st = conn.createStatement();
                         ResultSet rs = st.executeQuery(sql)) {
                        ResultSetMetaData md = rs.getMetaData();
                        int cols = md.getColumnCount();
                        StringBuilder sb = new StringBuilder("[");
                        boolean first = true;
                        while (rs.next()) {
                            if (!first) sb.append(",");
                            first = false;
                            sb.append("{");
                            for (int i = 1; i <= cols; i++) {
                                if (i > 1) sb.append(",");
                                String name = md.getColumnLabel(i);
                                String val = rs.getString(i);
                                sb.append("\"").append(name).append("\":\"")
                                  .append(val == null ? "" : val.replace("\\", "\\\\").replace("\"", "\\\"").replace("\n", "\\n"))
                                  .append("\"");
                            }
                            sb.append("}");
                        }
                        sb.append("]");
                        System.out.println(sb.toString());
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("{\"ok\":false,\"error\":\"" +
                e.getMessage().replace("\\", "\\\\").replace("\"", "\\\"") + "\"}");
            System.exit(0);
        }
    }
}
