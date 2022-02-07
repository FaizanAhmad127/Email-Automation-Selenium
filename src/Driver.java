import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.By;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Vector;

public class Driver {
    public static void main(String args[]) throws IOException, InterruptedException {

        String baseUrl = "https://accounts.google.com/signin/v2/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin";
        System.setProperty("webdriver.chrome.driver","E:\\EmailSend\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.get(baseUrl);



        driver.findElement(By.id("identifierId")).sendKeys("chrisg@louthcallanrenewables.com");
        Thread.sleep(2000);
        driver.findElement(By.id("identifierNext")).click();
        Thread.sleep(2000);
        driver.findElement(By.name("password")).sendKeys("Working11!");
        Thread.sleep(2000);
        driver.findElement(By.id("passwordNext")).click();
        Thread.sleep(25000);




        String filename = "E:\\EmailSend\\applying3.xls";
        FileInputStream fis = null;

        try {

            fis = new FileInputStream(filename);
            HSSFWorkbook workbook = new HSSFWorkbook(fis);
            HSSFSheet sheet = workbook.getSheetAt(0);
            Iterator rowIter = sheet.rowIterator();
            int k=1;
            rowIter.next();
            while(k<=50){

                HSSFRow myRow = (HSSFRow) rowIter.next();
                Iterator cellIter = myRow.cellIterator();
                Vector<String> cellStoreVector=new Vector<String>();
                while(cellIter.hasNext()){
                    HSSFCell myCell = (HSSFCell) cellIter.next();
                    String cellvalue = myCell.getStringCellValue();
                    cellStoreVector.addElement(cellvalue);
                }
                String name = null;
                String address = null;
                String email = null;
                String subject=null;

                int i = 0;
                name = cellStoreVector.get(i);
                address = cellStoreVector.get(i+1);
                email=cellStoreVector.get(i+2);
                System.out.println(name);
                System.out.println(address);
                System.out.println(email);
                System.out.println("\n");

                subject= "Good Afternoon "+name+","+"\n My name is Chris Glueck. I am the Vice President of Commercial" +
                        " Development for Louth Callan. We are a national development firm. " +
                        "We offer to pay to lease unused commercial space. Your facility is in one of our target markets. " +
                        "Is it possible to set up a quick phone call sometime this week?\n" +
                        "Thanks";




                    driver.findElement(By.cssSelector(".aic .z0 div")).click();
                    Thread.sleep(1000);

                    driver.findElement(By.name("to")).sendKeys(email);
                    driver.findElement(By.name("subjectbox")).sendKeys("Re: "+address);
                    driver.findElement(By.cssSelector(".Am.Al.editable.LW-avf")).sendKeys(subject);
                    Thread.sleep(3000);
                    driver.findElement(By.cssSelector(".T-I.J-J5-Ji.aoO.v7.T-I-atl.L3")).click();
                    Thread.sleep(1000);

            k++;
            }




        } catch (IOException e) {

            e.printStackTrace();

        } finally {

            if (fis != null) {

                fis.close();

            }

        }

//      showExelData(sheetData);

    }


}



