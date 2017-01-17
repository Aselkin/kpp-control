import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

/**
 * Created by Асюша on 26.06.2016.
 */
public class InControl {
    public static void main(String[] args) throws IOException {

        FileInputStream fileList = new FileInputStream(new File("List.xlsx"));
        XSSFWorkbook wbList = new XSSFWorkbook(fileList);
        XSSFSheet sheetList = wbList.getSheetAt(0);

        Set<Path> fileDate = getFileList();

        int indexRow = 0;
        for (Path path : fileDate) {
            FileInputStream fileKPP = new FileInputStream(path.toFile());
            XSSFWorkbook wbKPP = new XSSFWorkbook(fileKPP);
            XSSFSheet sheetKPP = wbKPP.getSheetAt(0);

            for (int indexkpp = 2; indexkpp < sheetKPP.getLastRowNum(); ++indexkpp) {
                Row rowKPP = sheetKPP.getRow(indexkpp);

                Cell colF = rowKPP.getCell(1);
                Cell colI = rowKPP.getCell(2);
                Cell colO = rowKPP.getCell(3);
                Cell colT = rowKPP.getCell(8);


                for (int indexrow = 1; indexrow < sheetList.getLastRowNum(); ++indexrow) {

                    Row row = sheetList.getRow(indexrow);
                    //For each row, iterate through all the columns

                    Cell colFIO = row.getCell(0);
                    Cell colTime = row.getCell(1);

                    if (colFIO == null || colTime == null) continue;

                    String cellFIO = colFIO.getStringCellValue();

                    Integer mustTime = tomin(colTime.getStringCellValue());

                    if (colFIO.getStringCellValue().replaceAll(" ", "").toUpperCase().equals(
                            (colF.getStringCellValue() +
                                    colI.getStringCellValue() +
                                    colO.getStringCellValue()).toUpperCase())) {


                        String sheetName = getSheetName(rowKPP);
                        XSSFSheet monthSheet = wbList.getSheet(sheetName);
                        if (monthSheet == null) {
                            monthSheet = wbList.createSheet(sheetName);
                            indexRow = 0;

                            Row headRow  = monthSheet.createRow(0);
                            headRow.createCell(0).setCellValue("ФИО");
                            headRow.createCell(1).setCellValue("Время");
                            for ( int d = 2; d < 33; d++ )
                                headRow.createCell(d).setCellValue(d - 1);

                        }

                        Row monthRow = getExistRow(monthSheet, colFIO.getStringCellValue());
                        if (monthRow == null)
                            monthRow = monthSheet.createRow(++indexRow);

                        Cell monthCellFIO = monthRow.createCell(0);
                        monthCellFIO.setCellValue(colFIO.getStringCellValue());

                        Cell monthCellTime = monthRow.createCell(1);
                        monthCellTime.setCellValue(colTime.getStringCellValue());
                        Cell monthCellDif = monthRow.createCell(getKPPDay(rowKPP) + 1);
                        Integer difTime = null;
                        try {
                            difTime = tomin(colT.getStringCellValue()) - mustTime;
                        } catch (NumberFormatException e) {
                            break;

                        }
                        monthCellDif.setCellValue(difTime < 0 ? 0 : difTime);
                        break;
                    }
                }
            }
        }
        fileList.close();
        FileOutputStream out = new FileOutputStream(new File("List.xlsx"));
        wbList.write(out);
        out.close();
    }

    private static Row getExistRow(XSSFSheet monthSheet, String findFIO) {
        for (int indexrow = 1; indexrow < monthSheet.getLastRowNum(); ++indexrow) {
            Row row = monthSheet.getRow(indexrow);

            Cell colFIO = row.getCell(0);

            if (findFIO.equals(colFIO.getStringCellValue())) {
                return row;
            }

        }
        return null;
    }


    private static Integer tomin(String timeValue) {

        String[] arrayTime = timeValue.split(":");
        return Integer.valueOf(arrayTime[0]) * 60 + Integer.valueOf(arrayTime[1]);

    }

    private static String getSheetName(Row rowKPP) {
        Cell colF = rowKPP.getCell(0);
        return colF.getStringCellValue().substring(3);

    }

    private static Set<Path> getFileList() throws IOException {
        Set<Path> fileDate = new TreeSet<>();
        Files.walk(Paths.get(".")).forEach(filePath -> {
            if (Files.isRegularFile(filePath) && filePath.toFile().getAbsolutePath().contains("КПП")) {
                fileDate.add(filePath);
                System.out.println(filePath);
            }
        });
        return fileDate;
    }


    private static int getKPPDay(Row rowKPP) {
        Cell colF = rowKPP.getCell(0);
        return Integer.valueOf(colF.getStringCellValue().substring(0, 2));

    }
}
