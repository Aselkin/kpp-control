import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;

/**
 * Created by Асюша on 26.06.2016.
 */
public class InControl {
    public static void main(String[] args) throws IOException {

        FileInputStream fileList = new FileInputStream(new File("List.xlsx"));
        XSSFWorkbook wbList = new XSSFWorkbook(fileList);

        ArrayList<Path> fileDate = getFileList();

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

                XSSFSheet sheetList = wbList.getSheetAt(0);
                for (int indexrow = 1; indexrow < sheetList.getLastRowNum(); ++indexrow) {

                    Row row = sheetList.getRow(indexrow);
                    //For each row, iterate through all the columns

                    Cell colFIO = row.getCell(0);
                    Cell colTime = row.getCell(1);

                    if (colFIO == null || colTime == null) continue;

                    String cellFIO = colFIO.getStringCellValue();
                    System.out.println(cellFIO);

                    Integer mustTime = Integer.valueOf(colTime.getStringCellValue().replaceAll(":", ""));


                    if (colFIO.getStringCellValue().replaceAll(" ", "").toUpperCase().equals(
                            colF.getStringCellValue() +
                                    colI.getStringCellValue() +
                                    colO.getStringCellValue().toUpperCase())) {


                        XSSFSheet monthSheet = wbList.getSheet(getSheetName(path));
                        if (monthSheet == null)
                            monthSheet = wbList.createSheet(getSheetName(path));

                        Row monthRow = monthSheet.createRow(monthSheet.getLastRowNum());
                        Cell monthCellFIO = monthRow.createCell(0);
                        monthCellFIO.setCellValue(colFIO.getStringCellValue());
                        Cell monthCellTime = monthRow.createCell(1);
                        Cell monthCellDif = monthRow.createCell(getKPPDay(path) + 1);

                    }

                }


            }


        }
    }

    private static String getSheetName(Path path) {
        String fileName = path.toFile().getName();
        System.out.println(fileName.substring(fileName.length() - 10, fileName.length() - 5));
        return fileName.substring(fileName.length() - 10, fileName.length() - 5);

    }

    private static ArrayList<Path> getFileList() throws IOException {
        ArrayList<Path> fileDate = new ArrayList<>();
        Files.walk(Paths.get(".")).forEach(filePath -> {
            if (Files.isRegularFile(filePath) && filePath.toFile().getAbsolutePath().contains("КПП")) {
                fileDate.add(filePath);
                System.out.println(filePath);
            }
        });
        return fileDate;
    }


    private static int getKPPDay(Path path) {
        String fileName = path.toFile().getName();
        System.out.println(fileName.substring(fileName.length() - 13, fileName.length() - 10));
        return Integer.valueOf(fileName.substring(fileName.length() - 13, fileName.length() - 10));

    }
}
