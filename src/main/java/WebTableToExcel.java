import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class WebTableToExcel {

	public static void main(String[] args) {
		WebDriverManager.chromedriver().setup();

		WebDriver driver = new ChromeDriver();

		driver.get("https://www.w3schools.com/html/html_tables.asp");
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		driver.manage().window().maximize();

		int row = driver.findElements(By.xpath("//table[@id='customers']/tbody/tr")).size();

		int column = driver.findElements(By.xpath("//table[@id='customers']/tbody/tr[1]/th")).size();
		String value = null;
		for (int i = 1; i <= row; i++) {
			for (int j = 1; j <= column; j++) {
				if (i == 1) {
					value = driver.findElement(By.xpath("//table[@id='customers']/tbody/tr[" + i + "]/th[" + j + "]"))
							.getText();
					System.out.println(value);
					ExcelWriter.exportDataToExcel(i - 1, j - 1, value);
				} else {
					value = driver.findElement(By.xpath("//table[@id='customers']/tbody/tr[" + i + "]/td[" + j + "]"))
							.getText();
					System.out.println(value);
					ExcelWriter.exportDataToExcel(i - 1, j - 1, value);
				}
			}
		}
		
		System.out.println("WebTable data succesfully exported to Excel");
		
	}

}
