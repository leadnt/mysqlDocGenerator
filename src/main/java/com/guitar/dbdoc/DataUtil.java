package com.guitar.dbdoc;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author hxy
 */
public class DataUtil {
    
    public List<Table> queryTables(String url,String user,String pwd) throws Exception{
        Class.forName("com.mysql.jdbc.Driver");
        List<Table> tables;
        try (Connection con = DriverManager.getConnection(url, user, pwd)) {
            tables = queryTables(con);
        }
        return tables;
    }
    private List<Table> queryTables(Connection con) throws Exception {
        ArrayList<Table> tables = new ArrayList<Table>();
        Statement st = null;
        ResultSet rs = null;
        try {
            st = con.createStatement();
            rs = st.executeQuery("show table status where comment != 'VIEW';");
            while (rs.next()) {
                String tableName = rs.getString("Name");
                String comment = rs.getString("Comment");
                Table table = new Table();
                table.setName(tableName);
                table.setComment(comment);
                table.setColumns(queryColumns(con,tableName));
                tables.add(table);
            }
        } finally {
            if (rs != null) {
                rs.close();
            }
            if (st != null) {
                st.close();
            }
        }
        return tables;
    }
    

    private List<Column> queryColumns(Connection con,String tableName) throws Exception {
        ArrayList<Column> columns = new ArrayList<Column>();
        Statement st = null;
        ResultSet rs = null;
        try {
            st = con.createStatement();
            rs = st.executeQuery("show full columns from " + tableName);
            while (rs.next()) {
                Column column = new Column();
                column.setType(rs.getString("Type"));
                column.setName(rs.getString("Field"));
                column.setComment(rs.getString("Comment"));
                column.setEmpty(rs.getString("Null"));
                column.setKey(rs.getString("Key"));
                column.setDefaultValue(rs.getString("Default"));
                column.setExtra(rs.getString("Extra"));
                column.setComment(rs.getString("Comment"));

                columns.add(column);
            }
        } finally {
            if (rs != null) {
                rs.close();
            }
            if (st != null) {
                st.close();
            }
        }
        return columns;
    }
    
    
    
}
