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

public class WriteExcel
{
    public String fileName, date;
    public ReadExcel student;
    public int columns;
    public WritableWorkbook workbook;
    public WritableSheet sheet1;
    public WritableFont times12 = new WritableFont(WritableFont.TIMES, 12);
    public WritableCellFormat times = new WritableCellFormat(times12);
    public CellView cv = new CellView();
    public String[][] cells;
    public File file;
    public WorkbookSettings wbSettings;
    public WritableSheet excelSheet;
    
    public WriteExcel(String fileName, String[][] cells) throws IOException, WriteException
    {
        this.fileName = fileName;
        this.cells = cells;
        
        file = new File("xl1.xls");
        wbSettings = new WorkbookSettings();
    
        wbSettings.setLocale(new Locale("en", "EN"));
    
        workbook = Workbook.createWorkbook(file, wbSettings);
        workbook.createSheet("Sheet 1", 0);
        excelSheet = workbook.getSheet(0);
        
        times.setWrap(false);
        cv.setFormat(times);
        
        for (int c = 0; c < cells.length; c++)
        {
            for (int r = 0; r < cells[c].length; r++)
            {
                if (r >= 1 && c == 2)
                {
                    tallyPoints(excelSheet, r);
                }
                else if (r >= 1 && (c == 0 || c >= 3) && !cells[c][r].equals(""))   //if it's supposed to be a number
                {
                    addNumber(excelSheet, c, r, Integer.parseInt(cells[c][r]));  //"addNumber" doesn't allow me to open excel, changed to "addLabel"
                }
                else addLabel(excelSheet, c, r, cells[c][r]);
            }
        }
        
        for (int r = 0; r < cells[0].length; r++)
        {
            tallyPoints(excelSheet, r);
        }
        
        workbook.write();
        workbook.close();
    }
    
    public WriteExcel(String fileName, ReadExcel student, String id, String name) throws IOException, WriteException
    {
        this.fileName = fileName;
        this.student = student;
        
        file = new File("xl1.xls");
        wbSettings = new WorkbookSettings();
    
        wbSettings.setLocale(new Locale("en", "EN"));
    
        workbook = Workbook.createWorkbook(file, wbSettings);
        workbook.createSheet("Sheet 1", 0);
        excelSheet = workbook.getSheet(0);
        
        times.setWrap(false);
        cv.setFormat(times);
        
        for (int c = 0; c < cells.length; c++)
        {
            for (int r = 0; r < cells[c].length; r++)
            {
                if (r >= 1 && c == 2)
                {
                    tallyPoints(excelSheet, r);
                }
                else if (r >= 1 && (c == 0 || c >= 3) && !cells[c][r].equals(""))   //if it's supposed to be a number
                {
                    addNumber(excelSheet, c, r, Integer.parseInt(cells[c][r]));  //"addNumber" doesn't allow me to open excel, changed to "addLabel"
                }
                else addLabel(excelSheet, c, r, cells[c][r]);
            }
        }
        
        addNumber(excelSheet, 0, cells[0].length, Integer.parseInt(id));    //added
        addNumber(excelSheet, cells.length - 1, cells[0].length, 1);        //added
        addLabel(excelSheet, 1, cells[0].length, name);                     //added
        
        tallyPoints(excelSheet, cells[0].length);
        
        workbook.write();
        workbook.close();
        student.changeFileName("xl1.xls");
        student.doExcel();
    }
    
    /*public WriteExcel(String fileName, int columns, String date) throws IOException, WriteException
    {
        this.fileName = fileName;
        this.date = date;
        this.columns = columns;
        
        /*try
        {
            workbook = Workbook.createWorkbook(new File("xl1.xls"), student.getWorkbook());   //fix
            workbook.createSheet("Sheet 1", 0);
            sheet1 = workbook.getSheet(0);
            
            WritableCell cell;
            Label day = new Label(columns, 0, date);
            cell = (WritableCell)day;
            
            sheet1.addCell(cell);
            workbook.write();
            workbook.close();
        }
        catch (Exception e) {}*
        
        file = new File("xl1.xls");
        wbSettings = new WorkbookSettings();
    
        wbSettings.setLocale(new Locale("en", "EN"));
    
        workbook = Workbook.createWorkbook(file, wbSettings);
        workbook.createSheet("Sheet 1", 0);
        excelSheet = workbook.getSheet(0);
        
        times.setWrap(false);
        cv.setFormat(times);
        
        for (int c = 0; c < student.getColumns(); c++)
        {
            for (int r = 0; r < student.getRows(); r++)
            {
                if (c == student.getDayColumn() && r == student.getNameRow())
                {
                    addNumber(excelSheet, c, r, 1);
                }
                else if (r >= 1 && (c == 0 || c >= 2) && !student.getCells(c, r).equals(""))   //if it's supposed to be a number
                {
                    addNumber(excelSheet, c, r, Integer.parseInt(student.getCells(c, r)));  //"addNumber" doesn't allow me to open excel, changed to "addLabel"
                }
                else addLabel(excelSheet, c, r, student.getCells(c, r));
            }
        }
        addLabel(excelSheet, columns, 0, date);
    
        workbook.write();
        workbook.close();
        student.doExcel();
    }*/
    
    public WriteExcel(String fileName, ReadExcel student, boolean start) throws IOException, WriteException
    {
        this.fileName = fileName;
        this.student = student;
        
        /*try
        {
            workbook = Workbook.createWorkbook(new File("xl1.xls"), student.getWorkbook());
            sheet1 = workbook.createSheet("Sheet 1", 0);
            
            Label write = new Label(student.getDayColumn(), student.getNameRow(), "1");
            sheet1.addCell(write);
            workbook.write();
            student.doExcel();
            //workbook.close();
        }
        catch (Exception e) {}*/
        
        File file = new File("xl1.xls");
        WorkbookSettings wbSettings = new WorkbookSettings();
    
        wbSettings.setLocale(new Locale("en", "EN"));
    
        workbook = Workbook.createWorkbook(file, wbSettings);
        workbook.createSheet("Sheet 1", 0);
        WritableSheet excelSheet = workbook.getSheet(0);
        
        times.setWrap(false);
        cv.setFormat(times);
        
        for (int c = 0; c < student.getColumns(); c++)
        {
            for (int r = 0; r < student.getRows(); r++)
            {
                try
                {
                    if (r >= 1 && c == 2)
                    {
                        tallyPoints(excelSheet, r);
                    }
                    else if (c == student.getDayColumn() && r == student.getNameRow() && !start)
                    {
                        addNumber(excelSheet, c, r, 1);
                        tallyPoints(excelSheet, r);
                    }
                    else if (r >= 1 && (c == 0 || c >= 3) && !student.getCells(c, r).equals(""))   //if it's supposed to be a number
                    {
                        addNumber(excelSheet, c, r, Integer.parseInt(student.getCells(c, r)));
                    }
                    else addLabel(excelSheet, c, r, student.getCells(c, r));
                }
                catch (Exception e)
                {
                    addLabel(excelSheet, c, r, "");
                }
            }
        }
        
        workbook.write();
        workbook.close();
        student.changeFileName("xl1.xls");
        student.doExcel();
    }
    
    /*public void write() throws IOException, WriteException {
        //File file = new File(fileName);
        File file = new File("xl1.xls");
        WorkbookSettings wbSettings = new WorkbookSettings();
    
        wbSettings.setLocale(new Locale("en", "EN"));
    
        workbook = Workbook.createWorkbook(file, wbSettings);
        workbook.createSheet("Sheet 1", 0);
        WritableSheet excelSheet = workbook.getSheet(0);
        createContent(excelSheet);
    
        workbook.write();
        workbook.close();
    }*/

    private void createContent(WritableSheet sheet) throws WriteException, RowsExceededException {
        // Write a few number
        /*for (int i = 1; i < 10; i++) {
          // First column
          addNumber(sheet, 0, i, i + 10);
          // Second column
          addNumber(sheet, 1, i, i * i);
        }*/
        // Lets calculate the sum of it
        /*StringBuffer buf = new StringBuffer();
        buf.append("SUM(A2:A10)");
        Formula f = new Formula(0, 10, buf.toString());
        sheet.addCell(f);
        buf = new StringBuffer();
        buf.append("SUM(B2:B10)");
        f = new Formula(1, 10, buf.toString());
        sheet.addCell(f);*/
        
        
        
    
        // now a bit of text
        /*for (int i = 12; i < 20; i++) {
          // First column
          addLabel(sheet, 0, i, "Boring text " + i);
          // Second column
          addLabel(sheet, 1, i, "Another text");
        }*/
    }

    private void addCaption(WritableSheet sheet, int column, int row, String s) throws RowsExceededException, WriteException {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }

    private void addNumber(WritableSheet sheet, int column, int row, Integer integer) throws WriteException, RowsExceededException {
        Number number;
        number = new Number(column, row, integer, times);
        sheet.addCell(number);
    }

    private void addLabel(WritableSheet sheet, int column, int row, String s) throws WriteException, RowsExceededException {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }
    
    private void addDate(WritableSheet sheet, int column, int row, Integer integer) throws WriteException, RowsExceededException {
        
    }
    
    public void newDate(int columns, String date) throws IOException, WriteException, RowsExceededException
    {
        addLabel(excelSheet, columns, 0, date);
        workbook.write();
        workbook.close();
    }
    
    public void tallyPoints(WritableSheet sheet, int row) throws WriteException, RowsExceededException
    {
        if (row >= 1)
        {
            StringBuffer buf = new StringBuffer();
            buf.append("SUMIF(D" + (row + 1) + ":IV" + (row + 1) + ",\"1\")");   //=SUMIF(D2:IV2,"1")
            Formula f = new Formula(2, row, buf.toString());
            sheet.addCell(f);
        }
    }
    
    public void addNew(String name, int ID, int rows, ReadExcel student) throws IOException, WriteException//, RowsExceededException
    {
        addNumber(excelSheet, 1, rows + 1, (Integer)ID);
        addLabel(excelSheet, 2, rows + 1, name);
        workbook.write();
        workbook.close();
    }
}