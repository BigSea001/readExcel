
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 **/
public class ExcelReaderUtil {
    //excel2003扩展名
    private static final String EXCEL03_EXTENSION = ".xls";
    //excel2007扩展名
    private static final String EXCEL07_EXTENSION = ".xlsx";

    private static FileOutputStream fos;
    private static FileOutputStream fos1;

    /**
     * 每获取一条记录，即打印
     */
    static void sendRows(String filePath, String sheetName, int sheetIndex, int curRow, List<String> cellList) {
        StringBuilder oneLineSb = new StringBuilder();
        for (String cell : cellList) {
            oneLineSb.append(cell.trim());
            oneLineSb.append("==");
        }
        String oneLine = oneLineSb.toString();
        // 去除最后一个分隔符
        if (oneLine.endsWith("==")) {
            oneLine = oneLine.substring(0, oneLine.lastIndexOf("=="));
        }
        // TODO 需要根据具体解析来做
        // msg_function_not_finished^功能正在升级中，敬请期待^Features are being upgraded, so please expect
        // <string name="msg_verify_pin_failure">验证PIN失败</string>
        String[] split = oneLine.split("==");
        StringBuilder znSb = new StringBuilder();
        znSb.append("    <string name=\"");
        if (split.length>2) {
            znSb.append(split[0]);
            znSb.append("\">");
            znSb.append(split[1]);
            znSb.append("</string>");
            znSb.append("\n");

            byte[] znByte = znSb.toString().getBytes();

            try {
                fos.write(znByte,0,znByte.length);
                fos.flush();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        StringBuilder enSb = new StringBuilder();
        enSb.append("    <string name=\"");
        if (split.length>2) {
            enSb.append(split[0]);
            enSb.append("\">");
            enSb.append(split[2]);
            enSb.append("</string>");
            enSb.append("\n");

            byte[] enByte = enSb.toString().getBytes();
            try {
                fos1.write(enByte,0,enByte.length);
                fos1.flush();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static void readExcel(String fileName) throws Exception {
        int totalRows;
        if (fileName.endsWith(EXCEL03_EXTENSION)) { //处理excel2003文件
            ExcelXlsReader excelXls = new ExcelXlsReader();
            totalRows = excelXls.process(fileName);
        } else if (fileName.endsWith(EXCEL07_EXTENSION)) {//处理excel2007文件
            ExcelXlsxReaderWithDefaultHandler excelXlsxReader = new ExcelXlsxReaderWithDefaultHandler();
            totalRows = excelXlsxReader.process(fileName);
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
        }
        System.out.println("发送的总行数：" + totalRows);
        fos.write("</resources>".getBytes(),0,"</resources>".length());
        fos.flush();
        fos1.write("</resources>".getBytes(),0,"</resources>".length());
        fos1.flush();
    }

    /**
     * fos输出中文字符串
     * fos1输出英文字符串
     */
    public static void main(String[] args) throws Exception {
        fos = new FileOutputStream("D:\\zn_string2.xml",false);
        fos1 = new FileOutputStream("D:\\en_string2.xml",false);
        String path = "D:\\v2.xlsx";
        fos.write("<resources>\n".getBytes(),0,"<resources>\n".length());
        fos.flush();
        fos1.write("<resources>\n".getBytes(),0,"<resources>\n".length());
        fos1.flush();
        ExcelReaderUtil.readExcel(path);
    }
}
