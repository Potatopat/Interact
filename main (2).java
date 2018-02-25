public class main
{
    public static void main(String[] args)
    {
        ReadExcel e = new ReadExcel("book1.xls");
        e.setDate();
        boolean x = e.checkSignIn("31");
        boolean y = e.checkSignIn("32");
        ReadExcel e2 = new ReadExcel("book1");
        e2.setDate();
        boolean x2 = e2.checkSignIn("31");
        boolean y2 = e2.checkSignIn("32");
    }
}