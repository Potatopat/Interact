import javax.swing.*;
import java.util.*;

public class Panel
{
    public static boolean running = true;
    public static int ID = 0, num = 0,                                                                                                                                 a = -2468;
    public static String studentId = "";
    public static JOptionPane j = new JOptionPane();
    public static String readFileName = "xl1.xls", writeFileName = "xl1.xls", readText = "students.txt";
    public static ReadExcel student = new ReadExcel(readFileName);
    //public static ReadExcel doubleSignIn = /*new ReadExcel(writeFileName)*/ student;
    public static int SOMENUMBERHAHAHA = a;    
    public static String[] Questions = new String[4000];
    public static Calendar now = Calendar.getInstance();
    public static WriteExcel write;
    public static ReadFile read;
    
    public static void main(String[] args)
    {
        try
        {
            write = new WriteExcel(writeFileName, student, true);
        }
        catch (Exception e) {}
        enterId();
    }
    
    public static void enterId()
    {
        try
        {
            do
            {
                Questions[num] = j.showInputDialog("What is your ID number?");
                studentId = Questions[num];
                try{
                    ID = Integer.parseInt(studentId);
                }
                catch(NumberFormatException nFE){
                    ID = 0;
                }
                if (ID == SOMENUMBERHAHAHA)
                    System.exit(0);
    
                /*try
                {
                    doubleSignIn.checkID(ID);
                }
                catch (Exception e){ }
    
                try
                {
                    student.checkID(ID);
                }
                catch (Exception e){ }*/
                //}
                
                try
                {
                    if (/*doubleSignIn*/student.checkSignIn(studentId) /*&& !(studentId == null || studentId.equals(""))*/)
                    {
                        JOptionPane.showMessageDialog(null, "You Already Signed in Today, " + student.returnName() + "!\nYou Have " + student.returnPoints(studentId) + " Points as of Today.\n\n© Pat Bartman  2014", "Oops!", JOptionPane.ERROR_MESSAGE);
                    }
                    else if (studentId.equals("") || studentId == null)
                    {
                        JOptionPane.showMessageDialog(null, "No ID Entered.\n\n© Pat Bartman  2014", "Oops!", JOptionPane.ERROR_MESSAGE);
                    }
                    else if (!student.checkID(studentId) /*&& !(studentId == null || studentId.equals(""))*/)
                    {
                        int option = JOptionPane.showConfirmDialog(null, "Not Valid, Create a New User?\n\n© Pat Bartman  2014", "Oops!", JOptionPane.YES_NO_OPTION);
                        /*if (option == 0)      //remove comments to fix
                        {
                            read = new ReadFile(readText, ID);
                            write.addNew(read.returnName(), ID, student.getRows(), student);
                            break;
                        }*/                     //remove comments to fix
                    }
                }
                catch (Exception e)
                {
                    if (e.getMessage() == null)     //check if cancel button is hit
                    {
                        JOptionPane.showMessageDialog(null, "You Can't Do That, Son!\n\n© Pat Bartman  2014", "No!", JOptionPane.ERROR_MESSAGE);
                    }
                    else
                    {
                        JOptionPane.showMessageDialog(null, "The Following Error Has Occurred:\n" + e + "\n\nPlease Try Again.\nIf the Error Continues to Occur, Contact a Superior.\n\n© Pat Bartman  2014", "No!", JOptionPane.ERROR_MESSAGE);
                    }
                }
            }while(!student.checkID(studentId) || /*doubleSignIn*/student.checkSignIn(studentId) || studentId == null || studentId.equals(""));
            //VotingGUI voting = new VotingGUI(student);
            
            student.addPoint(studentId);
            JOptionPane.showMessageDialog(null, "Welcome " + student.returnName() + "!\nYou Have " + student.returnPoints(studentId) + " Points as of Today.\n\n© Pat Bartman  2014", "Welcome!", JOptionPane.INFORMATION_MESSAGE);     //make centered
            
            num++;
            try
            {
                write = new WriteExcel(writeFileName, student, false);
            }
            catch (Exception e) {}
        }
        catch (Exception e) {}
        
        enterId();
    }
}
