package com.vpbank.sqlex;


import com.vpbank.sqlex.excel.*;
import org.apache.commons.cli.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.*;
import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class Main {

    final static Map<String, CellData> DATA_MAPS = new HashMap<>();
    final static Pattern EMPTY = Pattern.compile("\\s*");
    final static String CONFIG_FILE = "config.conf";
    final static String QUERY_FILE = "query.sql";

    static {
        DATA_MAPS.put("java.lang.Integer", NumberCellData.INSTANCE);
        DATA_MAPS.put("java.lang.Short", NumberCellData.INSTANCE);
        DATA_MAPS.put("java.lang.Boolean", BooleanCellData.INSTANCE);
        DATA_MAPS.put("java.lang.String", StringCellData.INSTANCE);
        DATA_MAPS.put("java.lang.Double", NumberCellData.INSTANCE);
        DATA_MAPS.put("java.lang.Byte", NumberCellData.INSTANCE);
        DATA_MAPS.put("java.lang.Long", NumberCellData.INSTANCE);
        DATA_MAPS.put("java.lang.Number", NumberCellData.INSTANCE);
        DATA_MAPS.put("java.math.BigDecimal", NumberCellData.INSTANCE);
        DATA_MAPS.put("java.util.Date", DateCellData.INSTANCE);
        DATA_MAPS.put("java.sql.Date", DateCellData.INSTANCE);
        DATA_MAPS.put("java.sql.Timestamp", DateCellData.INSTANCE);
    }

    public static void main(String[] args) throws ClassNotFoundException, IOException, SQLException {
        Options options = new Options();

        options.addOption(Option.builder("c").longOpt("config").hasArg(true).required(false).desc("Input config file").build());
        options.addOption(Option.builder("q").longOpt("query").hasArg(true).required(false).desc("Input select query").build());
        options.addOption(Option.builder("f").longOpt("file").hasArg(true).required(false).desc("Input from file that contains queries").build());
        options.addOption(Option.builder("e").longOpt("export").hasArg(true).required(false).desc("Export result to excel file").build());

        CommandLineParser parser = new DefaultParser();
        HelpFormatter formatter = new HelpFormatter();
        CommandLine cmd;
        try {
            cmd = parser.parse(options, args);
        } catch (ParseException e) {
            System.out.println("Error when parse arguments.");
            formatter.printHelp("utility-name", options);
            System.exit(1);
            return;
        }

        String configOptValue = cmd.getOptionValue('c');
        String queryOptValue = cmd.getOptionValue('q');
        String fileOptValue = cmd.getOptionValue('f');
        String exportOptValue = cmd.getOptionValue('e');
        boolean isExport = cmd.hasOption("e");
        String exportFile = isExport
                ? (isBlank(exportOptValue) ? "export-" + System.currentTimeMillis() + ".xlsx"
                : exportOptValue + (exportOptValue.toLowerCase().endsWith(".xlsx") ? "" : ".xlsx")) : null;

        String configFile = isBlank(configOptValue) ? CONFIG_FILE : configOptValue;

        if (!new File(configFile).exists()) {
            System.out.println("Config file is not exist.");
            System.exit(1);
            return;
        }

        List<String> queries = null;
        if (!isBlank(fileOptValue)) {
            if (!new File(fileOptValue).exists()) {
                System.out.println("Query file is not exist.");
                return;
            }
            queries = readQueries(fileOptValue);
        }

        if (!isBlank(queryOptValue)) {
            if (queries == null) {
                queries = new ArrayList<>();
            }
            queries.add(queryOptValue);
        }

        if (queries == null || queries.isEmpty()) {
            System.out.println("Not found any query.");
            System.exit(1);
            return;
        }

        Properties config = readConfig(configFile);

        String connectionUrl = config.getProperty("url");

        if (connectionUrl.contains("mysql")) {
            Class.forName("com.mysql.jdbc.Driver");
        } else {
            Class.forName("oracle.jdbc.driver.OracleDriver");
        }

        Properties properties = new Properties();
        properties.setProperty("user", config.getProperty("username"));
        properties.setProperty("password", config.getProperty("password"));

        try (Connection conn = DriverManager.getConnection(connectionUrl, properties)) {
            conn.setAutoCommit(false);
            try {
                if (isExport) {
                    try (OutputStream outputStream = new FileOutputStream(exportFile)) {
                        exportExcel(conn, queries, outputStream);
                    }
                } else {
                    showInConsole(conn, queries);
                }
                conn.commit();
            } catch (SQLException ex) {
                ex.printStackTrace();
                conn.rollback();
            }
        }
    }

    private static void exportExcel(Connection conn, List<String> queries, OutputStream outputStream) throws SQLException, IOException {
        Workbook workbook = new SXSSFWorkbook(100);

        CellStyle dateCellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-mm-dd hh:mm AM/PM"));

        int counter = 1;
        ResultSet resultSet;
        for (String query : queries) {
            System.out.println("Executing : " + query);
            try (Statement statement = conn.createStatement(ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY)) {
                if (statement.execute(query)) {
                    Sheet sheet = workbook.createSheet("Sheet" + (counter++));
                    resultSet = statement.getResultSet();
                    int total = writeToSheet(sheet, resultSet, dateCellStyle);
                    System.out.println("Total rows : " + total);
                } else {
                    int updateCount = statement.getUpdateCount();
                    System.out.println("Row updated: " + updateCount);
                }
            }
        }
        workbook.write(outputStream);
    }

    private static void showInConsole(Connection conn, List<String> queries) throws SQLException, IOException {
        for (String query : queries) {
            System.out.println("Executing : " + query);
            try (Statement statement = conn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY)) {
                if (statement.execute(query)) {
                    ResultSet resultSet = statement.getResultSet();
                    resultSet.last();
                    System.out.println("Total rows : " + resultSet.getRow());
                } else {
                    int updateCount = statement.getUpdateCount();
                    System.out.println("Row updated: " + updateCount);
                }
            }
        }
    }

    private static int writeToSheet(Sheet sheet, ResultSet resultSet, CellStyle dateCellStyle) throws IOException, SQLException {
        int total = 0;
        int rowIndex = 0;
        ResultSetMetaData metaData = resultSet.getMetaData();
        int columnCount = metaData.getColumnCount();

        Row labelRow = sheet.createRow(rowIndex++);
        for (int i = 1; i <= columnCount; i++) {
            Cell cell = labelRow.createCell(i - 1);
            cell.setCellValue(metaData.getColumnLabel(i));
        }

        Cell cell;
        CellData cellData;
        while (resultSet.next()) {
            total++;
            Row row = sheet.createRow(rowIndex++);
            for (int i = 1; i <= columnCount; i++) {
                try {
                    cell = row.createCell(i - 1);
                    cellData = DATA_MAPS.get(metaData.getColumnClassName(i));
                    cellData.setCellValue(cell, resultSet.getObject(i));

                    if (cellData instanceof DateCellData) {
                        cell.setCellStyle(dateCellStyle);
                    }

                } catch (NullPointerException ex) {
                    System.out.println("Error null : " + rowIndex + " " + i + " " + metaData.getColumnLabel(i) + " " + metaData.getColumnClassName(i));
                    throw ex;
                }
            }
        }

        return total;
    }

    private final static Properties readConfig(String configFile) {
        Properties properties = new Properties();
        try (InputStream inputStream = new FileInputStream(configFile)) {
            properties.load(inputStream);
            return properties;
        } catch (FileNotFoundException ex) {
            System.out.println("Config file not exist.");
        } catch (IOException ex) {
            System.out.println("Cannot read config file.");
        }
        return null;
    }

    private final static List<String> readQueries(String filePath) throws IOException {
        byte[] data = Files.readAllBytes(Paths.get(filePath));
        String content = new String(data, Charset.forName("UTF-8"));
        String[] queries = content.split("-- END --");
        return Arrays.stream(queries)
                .filter(it -> !EMPTY.matcher(it).matches())
                .map(it -> it.trim())
                .collect(Collectors.toList());
    }

    public static boolean isBlank(final CharSequence cs) {
        int strLen;
        if (cs == null || (strLen = cs.length()) == 0) {
            return true;
        }
        for (int i = 0; i < strLen; i++) {
            if (!Character.isWhitespace(cs.charAt(i))) {
                return false;
            }
        }
        return true;
    }
}
