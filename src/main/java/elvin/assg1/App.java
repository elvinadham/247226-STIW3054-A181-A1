package elvin.assg1;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


/**
 * Hello world!
 *
 */
public class App {

    private static List<public_file> dataArray = new ArrayList();
    private static String URL = "https://ms.wikipedia.org/wiki/Malaysia";

    public static void retrieveData() {
        try {

            Document document = Jsoup.connect(URL).get();

            for (Element row : document.select("table.wikitable:nth-of-type(4) tr")) {

                final String content = row.select("th").text();
                final String contentz = row.select("td").text();

                System.out.println(content + " = " + contentz);

                    Elements data1 = row.select("th");
                    Elements data2 = row.select("td");

                    String column1 = data1.text();
                    String column2 = data2.text();

                    dataArray.add(new public_file(column1, column2));
                }


        }catch (IOException e){
            System.out.println("Disconnected from main server:"+ URL);
        }
    }

    public static void ConvertToExcel(){

        if (dataArray.isEmpty()) {
            System.out.println("Test Failed");
            System.exit(0);
        }
        String excelFile = "Testing file.xlsx";

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Output example ");

        try {

            for (int i = 0; i < dataArray.size(); i++) {
                XSSFRow row = sheet.createRow(i);

                XSSFCell cell1 = row.createCell(0);
                cell1.setCellValue(dataArray.get(i).getColumn1());

                XSSFCell cell2 = row.createCell(1);
                cell2.setCellValue(dataArray.get(i).getColumn2());

            }
            FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
            workbook.write(fileOutputStream);
            fileOutputStream.flush();
            fileOutputStream.close();
            System.out.println("\n"+excelFile + "Transformed to excel.");
        } catch (IOException e) {
            System.out.println("\nERROR : Failed to write the file!");
        }
    }

    public static void main(String[] args) {
        retrieveData();
        ConvertToExcel();
    }
}


