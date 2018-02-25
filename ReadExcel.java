import java.io.*;
import java.util.*;
import java.text.*;
import java.lang.*;
import jxl.*;

public class ReadExcel
{
    public String fileName;
    public Workbook workbook;
    public Sheet sheet1;
    public int columns, rows, dayColumn = -1;
    public Cell[][] cell;
    public String[][] cells;
    public FormulaCell[] formula;
    public Calendar today = Calendar.getInstance();
    public int IDColumn = 0;    //change to coencide with excel table layout
    public String name;
    public int nameRow;
    public WriteExcel writeXl;
    public boolean first = true;
    public Cell[] temp;
    public String[] temps;
    
    public ReadExcel(String fileName)
    {
        this.fileName = fileName;
        doExcel();
        try
        {
            writeXl = new WriteExcel(fileName, cells);
        }
        catch (Exception e) {}
        setDate();
    }
    
    public void doExcel()
    {
        try
        {
            workbook = Workbook.getWorkbook(new File(fileName));
            sheet1 = workbook.getSheet(0);
            columns = sheet1.getColumns();  //can't have any blanks
            rows = sheet1.getRows();        //can't have any blanks
            cell = new Cell[columns][rows];
            cells = new String[columns][rows];
            for (int c = 0; c < columns; c++)
            {
                for (int r = 0; r < rows; r++)
                {
                    if (c == 2 && !first)
                    {
                        cell[c][r] = temp[r];
                        cells[c][r] = temps[r];
                    }
                    else
                    {
                        cell[c][r] = sheet1.getCell(c,r);
                        cells[c][r] = cell[c][r].getContents();
                    }
                }
            }
            
            if (first)
            {
                temp = cell[2];
                temps = cells[2];
            }
            first = false;
            
            workbook.close();   //will it affect the remainder of the program?
        }
        catch (Exception e){}
    }
    
    public void doExcel(boolean x)
    {
        try
        {
            workbook = Workbook.getWorkbook(new File(fileName));
            sheet1 = workbook.getSheet(0);
            cell = new Cell[columns][rows];
            for (int c = 0; c < columns; c++)
            {
                for (int r = 0; r < rows; r++)
                {
                    cell[c][r] = sheet1.getCell(c,r);
                    cells[c][r] = cell[c][r].getContents();
                }
            }
            
            workbook.close();   //will it affect the remainder of the program?
        }
        catch (Exception e){}
    }
    
    public String getFileName()
    {
        return fileName;
    }
    
    public Workbook getWorkbook()
    {
        return workbook;
    }
    
    public Sheet getSheet1()
    {
        return sheet1;
    }
    
    public Sheet getSheet(int num)
    {
        return workbook.getSheet(num);
    }
    
    public int getColumns()
    {
        return columns;
    }
    
    public Cell[] getColumn(int num)
    {
        return sheet1.getColumn(num);
    }
    
    public int getRows()
    {
        return rows;
    }
    
    public Cell[] getRow(int num)
    {
        return sheet1.getRow(num);
    }
    
    public Cell[][] getCells()
    {
        return cell;
    }
    
    public Cell getCell(int c, int r)
    {
        return cell[c][r];
    }
    
    public String getCells(int c, int r)
    {
        return cells[c][r];
    }
    
    public void setDate()   //make sure to call
    {
        int year = today.get(Calendar.YEAR);
        int month = today.get(Calendar.MONTH) + 1;
        int day = today.get(Calendar.DATE);
        
        for (int c = 0; c < columns; c++)
        {
            if (cells[c][0].equalsIgnoreCase(month + "/" + day + "/" + year))
            {
                dayColumn = c;
                c = columns;
            }
        }
        
        if (dayColumn == -1)
        {
            try                                                                 //recently changed this
            {
                String[][] temp = cells;
                cells = new String[columns + 1][rows];
                for (int c = 0; c < columns; c++)
                {
                    for (int r = 0; r < rows; r++)
                    {
                        cells[c][r] = temp[c][r];
                    }
                }
                dayColumn = columns;
                columns++;
                cells[dayColumn][0] = (month + "/" + day + "/" + year);
                writeXl.newDate(columns, (month + "/" + day + "/" + year));
                doExcel(true);    //put back when write works
            }
            catch (Exception e) {}
        }
    }
    
    public boolean checkSignIn(String id)
    {
        for (int r = 0; r < rows; r++)
        {
            if (cells[IDColumn][r].equalsIgnoreCase(id))
            {
                nameRow = r;
                name = cells[IDColumn + 1][nameRow];  //presuming that name column is after ID column
                try
                {
                    if (cells[dayColumn][nameRow].equalsIgnoreCase("1"))    //Won't work until write date works
                    {
                        return true;
                    }
                }
                catch (Exception ArrayIndexOutOfBoundsException) {}
                r = rows;
            }
        }
        
        return false;
    }
    
    public boolean checkID(String id)
    {
        for (int r = 0; r < rows; r++)
        {
            if (cells[IDColumn][r].equalsIgnoreCase(id))
            {
                name = cells[IDColumn + 1][r];  //presuming that name column is after ID column
                return true;
            }
        }
        
        return false;
    }
    
    public String returnName()
    {
        return name;
    }
    
    public int getDayColumn()
    {
        return dayColumn;
    }
    
    public int getNameRow()
    {
        return nameRow;
    }
    
    public void changeFileName(String fileName)
    {
        this.fileName = fileName;
    }
    
    public String returnPoints(String id)
    {
        for (int r = 0; r < rows; r++)
        {
            if (cells[IDColumn][r].equalsIgnoreCase(id))
            {
                return cells[2][r];
            }
        }
        return "-1";
    }
    
    public void addPoint(String id)     //make it also add to Cell type
    {
        for (int r = 0; r < rows; r++)
        {
            if (cells[IDColumn][r].equalsIgnoreCase(id))
            {
                int tem = Integer.parseInt(cells[2][r]);
                tem += 1;
                cells[2][r] = Integer.toString(tem);
            }
        }
        temps = cells[2];
    }
}