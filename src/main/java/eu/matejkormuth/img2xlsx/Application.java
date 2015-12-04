package eu.matejkormuth.img2xlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class Application {
    public Application(File img, File xlsx) throws IOException {
        short height = 16;
        short width = 16;

        BufferedImage image = ImageIO.read(img);

        //Get the workbook instance for XLS file
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Get first sheet from the workbook
        XSSFSheet sheet = workbook.createSheet();
        sheet.setTabColor(IndexedColors.RED.index);
        //sheet.setZoom();

        short size = 64;

        for (int x = 0; x < image.getWidth(); x++) {
            Row row = sheet.createRow(x);
            row.setHeight((short) (20 * 12));
            for (int y = 0; y < image.getHeight(); y++) {
                row.createCell(y);
            }
        }

        for (int y = 0; y < image.getHeight(); y++) {
            sheet.setColumnWidth(y, (int) (256 * 2.66f));
        }

        //Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = sheet.iterator();
        for (int x = 0; x < image.getWidth(); x++) {
            Row currentRow = rowIterator.next();
            // currentRow.setHeight(height);

            //Get iterator to all cells of current row
            Iterator<Cell> cellIterator = currentRow.cellIterator();
            for (int y = 0; y < image.getHeight(); y++) {
                Cell currentCell = cellIterator.next();

                XSSFCellStyle style = workbook.createCellStyle();

                Color color = new Color(image.getRGB(y, x));
                XSSFColor xssfColor = new XSSFColor(color);

                //System.out.println(xssfColor.getARGBHex());

                style.setFillForegroundColor(xssfColor);
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                currentCell.setCellStyle(style);
            }
        }

        FileOutputStream stream = new FileOutputStream(xlsx);
        workbook.write(stream);
        stream.close();
        workbook.close();
    }
}
