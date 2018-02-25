import java.util.*;
import java.io.File;
import java.io.IOException;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.WritableCell;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class WriteExcel3
{
    public String fileName;
    public Workbook origional;
    public ReadExcel read;
    
    private WritableCellFormat times;
    
    public WriteExcel3(String fileName, Workbook origional, ReadExcel read) throws IOException, WriteException
    {
        this.fileName = fileName;
        this.origional = origional;
        this.read = read;
        
        File file = new File(fileName);
        WorkbookSettings wbSettings = new WorkbookSettings();
        
        wbSettings.setLocale(new Locale("en", "US"));
        
        WritableWorkbook workbook = Workbook.createWorkbook(file, origional, wbSettings);
        workbook.createSheet("Sheet 1", 0);
        WritableSheet excelSheet = workbook.getSheet(0);
        
        WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
        times = new WritableCellFormat(times10pt);
        times.setWrap(false);
        
        CellView cv = new CellView();
        cv.setFormat(times);
        cv.setAutosize(true);
        
        for (int c = 0; c < read.getColumns(); c++)
        {
            for (int r = 0; r < read.getRows(); r++)
            {
                if (r >= 1 && (c == 0 || c >= 2) && !read.getCells(c, r).equals(""))   //if it's supposed to be a number
                {
                    addNumber(excelSheet, c, r, Integer.parseInt(read.getCells(c, r)));
                }
                else addLabel(excelSheet, c, r, read.getCells(c, r));
            }
        }
    }

    private void addNumber(WritableSheet sheet, int column, int row, Integer integer) throws WriteException, RowsExceededException
    {
        Number number = new Number(column, row, integer, times);
        sheet.addCell(number);
    }
    
    private void addLabel(WritableSheet sheet, int column, int row, String s) throws WriteException, RowsExceededException
    {
        Label label = new Label(column, row, s, times);
        sheet.addCell(label);
    }
}
