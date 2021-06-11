package handsi;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Shippingapache {
	public static void main(String[] args) throws IOException {
		System.setProperty("webdriver.chrome.driver","C:\\browser drivers\\chromedriver.exe");
		WebDriver driver =new ChromeDriver();
		driver.get("https://webapps.tekstac.com/shippingDetails/");
		driver.findElement(By.xpath("//a[contains(text(),'6543217')]")).click();
		String[] a=new String[30];
		int k=1;
		for(int i = 1;i<=12;i++)
		{
			for(int j=1;j<=2;j++)
			{
				try
				{
				String b=driver.findElement(By.xpath("//body[1]/div[2]/table[1]/tbody[1]/tr["+i+"]/td["+j+"]")).getText();
				a[k]=b;
				
				}
				catch (Exception e) {
					a[k]="";
				}
				k++;
			}
		}
		for(int i=1;i<=24;i++)
		{
			System.out.println(a[i]);
		}
	
	///File f=new File("");
		File src =new File("C:\\Users\\jp\\OneDrive\\Desktop\\pract1.xlsx");
		FileInputStream fil =new FileInputStream(src);
	XSSFWorkbook w=new XSSFWorkbook(fil);
	XSSFSheet s=w.getSheetAt(0);
	int z=1;
	for(int i=1;i<=12;i++)
	{
		for(int j=1;j<=2;j++)
		{
			
			s.getRow(i).createCell(j).setCellValue(a[z]);
			z++;
		}
	}
	
	FileOutputStream fout=new FileOutputStream(src);
	w.write(fout);
	w.close();
}
}
