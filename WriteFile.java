import java.io.*;
import java.util.*;

//import org.apache.poi.*;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.hssf.usermodel.*;

//import com.aspose.cells.*;

public class WriteFile
{
    BufferedWriter out;
    String read;
    int linenum = 3;
    
    public WriteFile(String textFileName, ReadFile student)
    {
        Calendar now = Calendar.getInstance();
        
        try
        {
            student.findId(student.returnID());
            out = new BufferedWriter(new FileWriter(textFileName, true));
            String phrase = "" + student.returnName() + "\t" + student.returnID() + "\t" + student.returnYear();
            for (int x = 0; x < answers.length; x++)
            {
                if (answers[x] != null)
                {
                    phrase += "\t" + answers[x];
                }
            }
            
            out.write(phrase);
            out.newLine();
            out.close();
        }catch(IOException e){
            System.out.println("There was a problem:" + e);
        }
        
        
        try
        {
            Workbook workbook = new Workbook(textFileName);
        }
        catch (Exception e){}
    }    
}