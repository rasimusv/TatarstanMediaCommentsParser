import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.jsoup.*;
import org.jsoup.nodes.*;
import org.jsoup.select.*;
import java.io.*;
import java.util.*;

public class Main {

    private static final String filepath = "/Users/rasimusv/Desktop/";
    private static XSSFWorkbook workbook;
    private static XSSFSheet sheet;
    private static int numOfRows;

    public static void main(String[] args) throws IOException {
        Scanner sc = new Scanner(System.in);
        initTable();
        System.out.println(TextColor.BLUE + "Hello! Enter the link!" + TextColor.RESET);
        boolean flag = true;
        while(flag) {
            String url = sc.next();
            runScript(url);
            flag = ConsoleQuestion.getAnswer("Do you want to continue?");
            if (flag) {
                System.out.println(TextColor.BLUE + "Enter the link!" + TextColor.RESET);
            }
        }
        write();
        renameFiles();
    }

    private static void renameFiles() {

        new File(filepath + "2.xlsx").renameTo(new File(filepath + "3.xlsx"));
        new File(filepath + "1.xlsx").renameTo(new File(filepath + "2.xlsx"));
        new File(filepath + "3.xlsx").renameTo(new File(filepath + "1.xlsx"));    }


    private static void runScript(String url) throws IOException {
        String sitename = url.split("//")[1].split("\\.")[0];
        switch (sitename) {
            case "vk" :
                VKPostParse(initSite(url), url);
                break;
            case "business-gazeta" :
            case "kam" :
                BOParse(initSite(url), url);
                break;
        }
    }

    private static void VKPostParse(Document site, String url) {
        Elements postAuthors = site.getElementsByAttributeValue("class", "post_author");

        String group = postAuthors
                .first()
                .child(0)
                .toString()
                .split(">")[1]
                .replaceAll("</a", "");

        Elements authors = site.getElementsByAttributeValue("class", "reply_author");
        Elements texts = site.getElementsByAttributeValue("class", "reply_text");

        ArrayList<Comment> list = new ArrayList<>();

        for (int i = 0; i < Math.min(authors.size(), texts.size()); i++){

            String author = authors.get(i)
                    .toString()
                    .split(">")[2]
                    .replace("</a", "");

            String text = texts.get(i)
                    .toString()
                    .replaceAll("<\\s*[^>]*>", "")
                    .replaceAll("\n", "");

            list.add(new Comment(text, author));
        }

        pasteData("ВК", group, list, url);
    }

    private static void BOParse (Document site, String url) {
        Elements elements = site.getElementsByAttributeValue("class", "comments-box__best");

        ArrayList<Comment> list = new ArrayList<>();

        String [] bodies = elements.toString().split("<p>");

        for (int i = 0; i < bodies.length - 1; i++) {
            list.add(new Comment(bodies[i + 1].split("</p>")[0].replaceAll("<br>", "\n"), ""));
        }

        pasteData("СМИ","Business Online", list, url);
    }

    private static void write() throws IOException {
        FileOutputStream out = new FileOutputStream(filepath + "2.xlsx");
        workbook.write(out);
        out.close();
    }

    private static void initTable() throws IOException {
        workbook = new XSSFWorkbook(filepath + "1.xlsx");
        sheet = workbook.getSheet("Лист1");
        numOfRows = sheet.getPhysicalNumberOfRows();
    }

    private static Document initSite(String url) throws IOException {
        return Jsoup.connect(url).get();
    }

    private static void pasteData(String source, String group, ArrayList<Comment> list, String url) {

        CellStyle defaultStyle = sheet.getRow(1).getCell(1).getCellStyle();
        CellStyle hyperLinkStyle = sheet.getRow(1).getCell(5).getCellStyle();

        for (int i = 0; i < list.size(); i++) {
            XSSFRow r = sheet.createRow(numOfRows + i);
            XSSFCell [] ca = new XSSFCell[6];
            for (int j = 0; j < 6; j++) {
                ca[j] = r.createCell(j);
                ca[j].setCellStyle(defaultStyle);
            }
            ca[0].setCellValue(source);
            ca[1].setCellValue(group);
            ca[2].setCellValue(list.get(i).author);
            ca[3].setCellValue(list.get(i).text);
            ca[5].setCellValue(url);

            Hyperlink href = workbook
                    .getCreationHelper()
                    .createHyperlink(HyperlinkType.URL);

            href.setAddress(url);
            ca[5].setHyperlink(href);
            ca[5].setCellStyle(hyperLinkStyle);
        }
        numOfRows = sheet.getPhysicalNumberOfRows();
        System.out.println(numOfRows);
    }
}