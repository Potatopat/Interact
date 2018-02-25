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

public class WriteExcel4
{
    public String fileName;
    
    public WriteExcel4 (String fileName, String[][] cells) throws IOException, WriteException
    {
        this.fileName = fileName.substring(0, fileName.length() - 4) + ".xls";
    }
}