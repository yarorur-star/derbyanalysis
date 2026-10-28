// ExportToExcel.java — выгрузка table1..table10.sql из Derby в один Excel (лист на таблицу).
// Сборка/запуск (Derby embedded + Apache POI):
//   javac -cp "derby.jar;poi-5.2.5.jar;poi-ooxml-5.2.5.jar" ExportToExcel.java
//   java  -cp ".;derby.jar;poi-5.2.5.jar;poi-ooxml-5.2.5.jar;<зависимости poi>" \
//         -Dderby.system.home=C:\path\to\dbhome ExportToExcel C:\path\to\db\FISHSTAT report.xlsx sql
// Требования: derby.jar, Apache POI (poi, poi-ooxml, xmlbeans, commons-compress и т.д.)
import java.io.*;
import java.nio.file.*;
import java.sql.*;
import java.util.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

public class ExportToExcel {
    static final int[] TABLES = {1,2,3,4,5,6,7,8,9,10};
    static final String[] SHEETS = {
        "1_TopCountries","2_GroupsCommodities","3_CountriesYears",
        "4_Group1_Fresh","5_Group2_Dried","6_Group3_Pelagic","7_Group4_Demersal",
        "8_Group5_Crust","9_Group6_Mollusc","10_Group7_Other"};
    static final Set<String> DROP = new HashSet<>(Arrays.asList("LVL","GRPKEY","ORDKEY"));

    public static void main(String[] a) throws Exception {
        String dbPath = a[0];
        String out = a.length > 1 ? a[1] : "report_java.xlsx";
        String sqlDir = a.length > 2 ? a[2] : ".";
        try (XSSFWorkbook wb = new XSSFWorkbook();
             Connection cn = DriverManager.getConnection("jdbc:derby:" + dbPath + ";create=false")) {

            CellStyle hSt = wb.createCellStyle(); // шапка: жирный белый на тёмно-синем
            Font hF = wb.createFont(); hF.setBold(true); hF.setColor(IndexedColors.WHITE.getIndex());
            hSt.setFont(hF);
            hSt.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
            hSt.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle iSt = wb.createCellStyle(); // строки товаров: жирный, голубой фон
            Font iF = wb.createFont(); iF.setBold(true); iSt.setFont(iF);
            iSt.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
            iSt.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            DataFormat df = wb.createDataFormat();
            for (int k = 0; k < TABLES.length; k++) {
                int tno = TABLES[k];
                File f = new File(sqlDir, "table" + tno + ".sql");
                if (!f.exists()) continue;
                String sql = new String(Files.readAllBytes(f.toPath()), "UTF-8").trim();
                if (sql.endsWith(";")) sql = sql.substring(0, sql.length() - 1);
                System.out.println("Выполняю table" + tno + ".sql ...");
                try (Statement st = cn.createStatement(); ResultSet rs = st.executeQuery(sql)) {
                    ResultSetMetaData md = rs.getMetaData();
                    List<Integer> keep = new ArrayList<>();
                    for (int i = 1; i <= md.getColumnCount(); i++)
                        if (!DROP.contains(md.getColumnName(i).toUpperCase())) keep.add(i);
                    XSSFSheet sh = wb.createSheet(SHEETS[k]);
                    sh.createFreezePane(0, 1);
                    Row hr = sh.createRow(0);
                    for (int j = 0; j < keep.size(); j++) {
                        Cell c = hr.createCell(j);
                        c.setCellValue(md.getColumnLabel(keep.get(j)));
                        c.setCellStyle(hSt);
                    }
                    boolean items = tno >= 4;
                    int r = 1;
                    while (rs.next()) {
                        Row row = sh.createRow(r++);
                        String first = rs.getString(keep.get(0));
                        boolean isItem = items && (first == null || !first.startsWith("     - "));
                        for (int j = 0; j < keep.size(); j++) {
                            Object v = rs.getObject(keep.get(j));
                            Cell c = row.createCell(j);
                            if (v instanceof Number) {
                                c.setCellValue(((Number) v).doubleValue());
                                c.setDataFormat(df.getFormat("#,##0"));
                            } else {
                                c.setCellValue(v == null ? "" : v.toString());
                            }
                            if (isItem) c.setCellStyle(iSt);
                        }
                    }
                    for (int j = 0; j < keep.size(); j++) sh.autoSizeColumn(j);
                    System.out.println("  лист '" + SHEETS[k] + "': " + (r - 1) + " строк");
                }
            }
            try (FileOutputStream fo = new FileOutputStream(out)) { wb.write(fo); }
        }
        System.out.println("Готово: " + new File(out).getAbsolutePath());
    }
}
