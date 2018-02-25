import java.util.Scanner;
import java.io.*;
import java.text.DecimalFormat;

public class ReadFile
{
    private Scanner fileScan, lineScan;
    private String fileName;
    private Scanner scan = new Scanner (System.in);
    private String line, firstName, lastName, name;
    private int textID;
    
    public ReadFile(String textFileName, int ID)
    {
        try
        {
            fileName = textFileName;
            
            fileScan = new Scanner (new FileReader(fileName));
            line = fileScan.nextLine();
            do
            {
                line = fileScan.nextLine();
                lineScan = new Scanner (line);
                lineScan.useDelimiter("\t");
                
                lastName = lineScan.next();
                firstName = lineScan.next();
                name = firstName + " " + lastName;
                textID = Integer.parseInt(lineScan.next());
            }
            while (ID != textID && fileScan.hasNext());
        }
        catch (Exception e) {}
    }
    
    public String returnName()
    {
        return name;
    }
    
    public int returnID()
    {
        return textID;
    }
}