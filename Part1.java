import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.List;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Part1 {
	 static Workbook workbook = new XSSFWorkbook();
	    static Sheet sheet = workbook.createSheet("AutomationSteps");
	    static int rowCount = 0;

	    // Initialize Excel table headers
	    public static void initializeHeaders() {
	        Row headerRow = sheet.createRow(rowCount++);
	        String[] headers = {"Step No", "Description", "Status", "Timestamp"};

	        for (int i = 0; i < headers.length; i++) {
	            Cell cell = headerRow.createCell(i);
	            cell.setCellValue(headers[i]);

	            CellStyle style = workbook.createCellStyle();
	            Font font = workbook.createFont();
	            font.setBold(true);
	            style.setFont(font);
	            cell.setCellStyle(style);

	            // Optional: Auto-size columns
	            sheet.autoSizeColumn(i);
	            
	           
	        }
	    }

	    public static void logStep(int stepNo, String Description, String status) {
	        Row row = sheet.createRow(rowCount++);

	        // Wrap style for Description column only
	        CellStyle wrapStyle = workbook.createCellStyle();
	        wrapStyle.setWrapText(true);

	        row.createCell(0).setCellValue(stepNo);

	        // Description cell with wrap text
	        Cell descCell = row.createCell(1);
	        descCell.setCellValue(Description);
	        descCell.setCellStyle(wrapStyle);

	        row.createCell(2).setCellValue(status);

	        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
	        row.createCell(3).setCellValue(timestamp);

	        // Auto-size only these columns
	        sheet.autoSizeColumn(0);
	        sheet.autoSizeColumn(2);
	        sheet.autoSizeColumn(3);
	        
	        
	        sheet.setColumnWidth(0, 3000);
	        sheet.setColumnWidth(2, 4000);
	        sheet.setColumnWidth(3, 5000);

	        // Keep Description column fixed and wrapped
	        sheet.setColumnWidth(1, 10000);
	    
	            
	    

	    }

	    public static void main(String[] args) {

	        initializeHeaders();
	        WebDriver driver = null;

	        try {
	            // Launch Chrome (visible)
	            ChromeOptions options = new ChromeOptions();
	            options.addArguments("--start-maximized");
	            driver = new ChromeDriver(options);

	            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	            Actions action = new Actions(driver);

	            int step = 1;

	            // Open browser and URL
	            driver.get("https://dev-orders.richervalues.com/auth/login");
	            logStep(step++, "Browser opened and URL loaded", "PASS");
	            Thread.sleep(2000);

	            // Login
	            WebElement username = driver.findElement(By.id("username"));
	            username.sendKeys("Test_Administrator");
	            logStep(step++, "Entered username", "PASS");
	            Thread.sleep(1000);

	            WebElement password = driver.findElement(By.id("password"));
	            password.sendKeys("Browse]1Pope|8Vivacious/6Battered+1Swiftness");
	            logStep(step++, "Entered password", "PASS");
	            Thread.sleep(1000);

	            WebElement signin = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/div/div[1]/div/form/div/div[3]/div/div/button"));
	            signin.click();
	            logStep(step++, "Clicked Sign In button", "PASS");
	            Thread.sleep(2000);

	            // Place New Order
	            WebElement placeNewOrder = wait.until(ExpectedConditions.elementToBeClickable(By.linkText("Place New Order")));
	            action.moveToElement(placeNewOrder).perform();
	            placeNewOrder.click();
	            logStep(step++, "Clicked on 'Place New Order'", "PASS");
	            Thread.sleep(1000);
	            
	            
	            try {
		            WebDriverWait wait14 = new WebDriverWait(driver, Duration.ofSeconds(15));

		            // Wait for the "Got it, thank you" button dynamically
		            WebElement gotItButton = wait14.until(ExpectedConditions.elementToBeClickable(
		                By.xpath("//button[contains(text(),'Got it')] | //button[contains(text(),'Got it, thank you')]")
		            ));

		            gotItButton.click();
		            System.out.println("✅ Clicked on 'Got it, thank you' dynamically!");

		            Thread.sleep(500);

		        } catch (org.openqa.selenium.TimeoutException e) {
		            System.out.println("ℹ️ 'Got it, thank you' popup not found — continuing normally.");
		        } catch (Exception e) {
		            System.out.println("⚠️ Unexpected error while closing popup: " + e.getMessage());
		        }

		        Thread.sleep(1000);

	            // Evaluations menu
	            WebElement evaluations = wait.until(ExpectedConditions.elementToBeClickable(By.id("eval")));
	            action.moveToElement(evaluations).perform();
	            evaluations.click();
	            logStep(step++, "Clicked on 'Evaluations'", "PASS");
	            Thread.sleep(1000);

	            // Client selection
	            WebElement selectClient = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("input.rbt-input-main")));
	            selectClient.click();
	            selectClient.clear();
	            selectClient.sendKeys("Anhas Client Company -Dev");
	            logStep(step++, "Typed 'Anhas Client Company-Dev' in client dropdown", "PASS");
	            Thread.sleep(500);

	            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(".rbt-menu")));
	            selectClient.sendKeys(Keys.ARROW_DOWN);
	            selectClient.sendKeys(Keys.ENTER);
	            logStep(step++, "Selected first client dropdown option", "PASS");
	            Thread.sleep(1000);

	            // Loan Number
	            WebElement loanNumber = driver.findElement(By.id("client_loan_number"));
	            loanNumber.sendKeys("LN-00098765");
	            logStep(step++, "Entered Loan Number: LN-00098765", "PASS");
	            Thread.sleep(2000);
                System.out.println(); 
	            
	            // Property selection
	            WebElement propertyVacant = driver.findElement(By.xpath("//*[@id='root']/div[1]/div[2]/div/div[2]/form/div[1]/div[3]/div/div/div[2]/label"));
	            propertyVacant.click();
	            logStep(step++, "Selected 'Property Vacant Land'", "PASS");
	            Thread.sleep(500);

	            WebElement propertyUndergoing = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/div[2]/form/div[1]/div[3]/div[2]/div/div[2]/label"));
	            propertyUndergoing.click();
	            logStep(step++, "Selected 'Property Currently Undergoing'", "PASS");
	            Thread.sleep(1000);
	            System.out.println();

	            // Report Type
	            WebElement reportType = driver.findElement(By.xpath("//*[@id='report_type_1']/div/label"));
	            reportType.click();
	            logStep(step++, "Selected 'Reno ARV' report type", "PASS");
	            Thread.sleep(1000);

//	            // Scroll & Next button
//	            ((JavascriptExecutor) driver).executeScript("window.scrollBy(0,500)");
//	            WebElement nextButton = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/div[2]/form/div[2]/div/div[2]/button/span[1]"));
//	            nextButton.click();
//	            logStep(step++, "Clicked on 'Next' button", "PASS");
//	            Thread.sleep(2000);
//	          

	           
	            JavascriptExecutor js = (JavascriptExecutor) driver;
	            js.executeScript("window.scrollBy(0, 500)");
	            logStep(step++, "Page scrolled down 500px", "PASS");
	            Thread.sleep(3000);
	            

	            // Inspection Needed
	            WebElement InspectionNeeded = driver.findElement(
	                    By.xpath("//label[@for='radio-inspection_type-none']"));
	            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", InspectionNeeded);
	            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", InspectionNeeded);
	            logStep(step++, "Clicked on Inspection Needed (None)", "PASS");
	  //          InspectionNeeded.click();
	            Thread.sleep(2000);
	            System.out.println();
	            
	            
	            
	        
	            

	            // Turnaround Time
	            WebDriverWait wait2 = new WebDriverWait(driver, Duration.ofSeconds(10));
	            WebElement TurnaroundTimeLabel = wait.until(
	                ExpectedConditions.elementToBeClickable(By.xpath("//label[@for='radio-turnaround_time-standard']"))
	            );
	            TurnaroundTimeLabel.click();
	            logStep(step++, "Selected standard Turnaround Time", "PASS");
	            Thread.sleep(3000);


	            // Closing Date
//	            WebElement ClosingDate = driver.findElement(By.name("select_date"));
//	            ClosingDate.click();
//	            logStep(step++, "Clicked on Closing Date calendar input", "PASS");
//	            Thread.sleep(1000);


	            //
	            try {
	                WebElement calendarInput = wait.until(
	                    ExpectedConditions.elementToBeClickable(By.name("select_date"))
	                );
	                calendarInput.click();
	                Thread.sleep(800);
	                logStep(step++, "Calendar opened successfully", "PASS");

	                LocalDate target = LocalDate.now().plusDays(2);

	                while (true) {
	                    WebElement monthHeader = driver.findElement(By.cssSelector(".react-datepicker__current-month"));
	                    String header = monthHeader.getText(); // Example: "January 2025"

	                    String monthName = target.getMonth().getDisplayName(TextStyle.FULL, Locale.ENGLISH);
	                    String yearStr = String.valueOf(target.getYear());

	                    if (header.contains(monthName) && header.contains(yearStr)) {
	                        break;
	                    }

	                    WebElement next = driver.findElement(By.cssSelector(".react-datepicker__navigation--next"));
	                    next.click();
	                    Thread.sleep(500);
	                }

	                // --- Select Day ---
	                int day = target.getDayOfMonth();

	                WebElement dayEl = wait.until(ExpectedConditions.elementToBeClickable(
	                    By.xpath("//div[contains(@class,'react-datepicker__day') and not(contains(@class,'outside-month')) and text()='" + day + "']")
	                ));

	                js.executeScript("arguments[0].click();", dayEl);

	                logStep(step++, "Selected future date: " + target, "PASS");

	            } catch (Exception e) {
	                logStep(step++, "Future date selection failed: " + e.getMessage(), "FAIL");
	            }

	            //
	            // Loan Officer Selection
	            //

	            WebDriverWait wait15 = new WebDriverWait(driver, Duration.ofSeconds(10));

	            try {
	                WebElement selectInput = wait15.until(ExpectedConditions.elementToBeClickable(
	                    By.cssSelector(".form-control-sm")
	                ));
	                selectInput.click();
	                selectInput.sendKeys("Test_2 Test_2");
	                logStep(step++, "Typed Loan Officer Name: Test_2 Test_2", "PASS");

	                WebElement option = wait15.until(ExpectedConditions.elementToBeClickable(
	                    By.xpath("//*[contains(text(),'Test_2 Test_2') and not(ancestor::option)]")
	                ));

	                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", option);
	                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", option);

	                logStep(step++, "Loan Officer option selected", "PASS");

	            } catch (Exception e) {
	                logStep(step++, "Failed to select Loan Officer: " + e.getMessage(), "FAIL");
	            }
                     Thread.sleep(500);
                     

	            // Lender Entity
	            WebElement LenderEntity = driver.findElement(By.id("company_entity_id"));
	            LenderEntity.click();
	            logStep(step++, "Opened Lender Entity dropdown", "PASS");
	            Thread.sleep(1000);

	            WebElement select = driver.findElement(By.xpath("//*[@id=\"company_entity_id\"]/option[2]"));
	            select.click();
	            logStep(step++, "Selected Lender Entity option", "PASS");
	            Thread.sleep(1000);


	            // Scroll again
	            js.executeScript("window.scrollBy(0, 500)");
	            logStep(step++, "Scrolled down 500px before Next", "PASS");
	            Thread.sleep(1000);


	            // Next Button
	            WebElement NextButton = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/div[2]/form/div[2]/div/div[2]/button/span[1]"));
	            NextButton.click();
	            logStep(step++, "Clicked Next Button", "PASS");
	            Thread.sleep(1000);
	            
//	          Step 2
	            
	         // Property Type
	            WebElement propertyType = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/form/div[1]/div[1]/div/div/fieldset/div[1]/div[1]/label"));
	            propertyType.click();
	            logStep(step++, "Selected property type: Single", "PASS");
	            Thread.sleep(1000);


	            // Address Input
	            WebElement addressInput = driver.findElement(By.className("pac-target-input"));
	            addressInput.sendKeys("2560 DUKELAND DR");
	            logStep(step++, "Entered address: 2560 DUKELAND DR", "PASS");

	            WebDriverWait wait123 = new WebDriverWait(driver, Duration.ofSeconds(10));

	            // First Address Suggestion
	            WebElement firstSuggestion = wait.until(
	                ExpectedConditions.visibilityOfElementLocated(By.className("pac-item")));
	            
	            firstSuggestion.click();
	            logStep(step++, "Selected first suggested address", "PASS");
	            Thread.sleep(1000);


	            // Property Type (Single-Family)
	            WebElement proType = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/form/div[1]/div[5]/div[2]/div/div/fieldset/div[1]/div[1]/label"));
	            proType.click();
	            logStep(step++, "Selected property subtype: Single-Family", "PASS");
	            Thread.sleep(500);


	            // Condition (Moderate)
	            WebElement condition = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/form/div[1]/div[5]/div[3]/div/table/tbody/tr[5]/td[2]"));
	            condition.click();
	            logStep(step++, "Selected property condition: Moderate", "PASS");
	            Thread.sleep(500);


	            // Above Grade SqFt
	            WebElement sqft = driver.findElement(By.id("above_grade_sqft"));
	            sqft.sendKeys("3000");
	            logStep(step++, "Entered Above Grade Sqft: 3000", "PASS");
	            Thread.sleep(500);


	            // Bedrooms
	            WebElement bedroom = driver.findElement(By.id("bedrooms"));
	            bedroom.sendKeys("3");
	            logStep(step++, "Entered Bedrooms: 3", "PASS");
	            Thread.sleep(500);


	            // Bathrooms
	            WebElement Bathroom = driver.findElement(By.id("bathrooms"));
	            Bathroom.sendKeys("2");
	            logStep(step++, "Entered Bathrooms: 2", "PASS");
	            Thread.sleep(500);


	            // Year Built
	            WebElement YearBuilt = driver.findElement(By.id("year_built"));
	            YearBuilt.sendKeys("1994");
	            logStep(step++, "Entered Year Built: 1994", "PASS");
	            Thread.sleep(500);


	            // Stories
	            WebElement Stories = driver.findElement(By.id("stories"));
	            Stories.sendKeys("2");
	            logStep(step++, "Entered Stories: 2", "PASS");
	            Thread.sleep(500);


	            // Lot Size
	            WebElement LotSize = driver.findElement(By.id("lot_size_square_feet"));
	            LotSize.sendKeys("20000");
	            logStep(step++, "Entered Lot Size: 20000 sqft", "PASS");
	            Thread.sleep(500);


	            // Garage Space
	            WebElement GarageSpace = driver.findElement(By.id("garage_spaces"));
	            GarageSpace.sendKeys("1");
	            logStep(step++, "Entered Garage Space: 1", "PASS");
	            Thread.sleep(500);
	            
	            // Renovation Budget
	            WebElement Budget = driver.findElement(By.id("borrower_budget"));
	            Budget.sendKeys("2500000");
	            logStep(step++, "Enter Renovation Budget", "Pass");
	            Thread.sleep(500);


	            // Valuation Report
	            WebElement Valutionreport = driver.findElement(By.id("valuation_commentary_or_instruction"));
	            Valutionreport.sendKeys("Additional comment");
	            logStep(step++, "Entered Valuation Report comment", "PASS");
	            Thread.sleep(500);


	            // Next Button
	            WebDriverWait wait4 = new WebDriverWait(driver, Duration.ofSeconds(10));
	            WebElement Nextbtn = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[contains(.,'Next')]")));
	            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", Nextbtn);
	            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", Nextbtn);
	            logStep(step++, "Clicked on Next button (Property Details)", "PASS");
	            Thread.sleep(500);
	            
	            // Fix later
		         
		         try {
		        	    WebDriverWait wait18 = new WebDriverWait(driver, Duration.ofSeconds(10));

		        	    // Dynamic locator based on button text "Fix Later"
		        	    WebElement fixLaterButton = wait18.until(ExpectedConditions.elementToBeClickable(
		        	        By.xpath("//button[contains(text(), 'Fix Later')]")
		        	    ));

		        	    
//		        	    Actions action = new Actions(driver);
		        	    action.moveToElement(fixLaterButton).perform();
		        	    System.out.println("✅ Mouse moved to 'Fix Later' button");

		        	    // Click the button
		        	    fixLaterButton.click();
		        	    System.out.println("✅ 'Fix Later' button clicked successfully");

		        	    Thread.sleep(500); 

		        	} catch (Exception e) {
		        	    System.out.println("⚠️ Could not click 'Fix Later' button: " + e.getMessage());
		        	}
	            
	            
//			       Step 3
		         
	            WebElement Nextbtn2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/form/div[2]/div/div[3]/button[2]"));
	            Nextbtn2.click();
	            logStep(step++, "Clicked on Next button (Final Step of this page)", "PASS");
	            Thread.sleep(500);
	            
//		       Step 4
		         
	            js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
	            logStep(step++, "Scrolled to bottom of the page", "PASS");

	            WebElement save = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/form/div[2]/div/div[3]/button"));
	            save.click();
	            logStep(step++, "Clicked on Save button", "PASS");

	            Thread.sleep(500);
	        
	    
		         
//		       Step 5
		         
		         
		         
		         
	            try {

	                // ============================
	                // PAYMENT BUTTON
	                // ============================
	                WebDriverWait wait5 = new WebDriverWait(driver, Duration.ofSeconds(10));

	                WebElement payment = wait5.until(
	                        ExpectedConditions.elementToBeClickable(By.cssSelector("button[id^='action_button_']"))
	                );

	                js.executeScript("arguments[0].scrollIntoView(true);", payment);
	                Thread.sleep(2000);
	                js.executeScript("arguments[0].click();", payment);
	                logStep(step++, "Clicked Payment button", "PASS");


	                // ============================
	                // SELECT CREDIT CARD OPTION
	                // ============================
	                WebDriverWait wait6 = new WebDriverWait(driver, Duration.ofSeconds(15));

	                WebElement Creditcard = wait6.until(ExpectedConditions.visibilityOfElementLocated(
	                        By.xpath("/html/body/div[4]/div/div/div[2]/div/div[2]/form/div[1]/div/div/div/div[2]/label/span")
	                ));

	                js.executeScript("arguments[0].scrollIntoView(true);", Creditcard);
	                Thread.sleep(3000);

	                Actions actions = new Actions(driver);
	                actions.moveToElement(Creditcard).pause(Duration.ofMillis(500)).click().perform();
	                logStep(step++, "Selected Credit Card option", "PASS");


	                // ============================
	                // CLICK ADD NEW SOURCE
	                // ============================
	                try {
	                    WebElement addNew = wait6.until(ExpectedConditions.elementToBeClickable(
	                            By.xpath("//label[contains(text(),'Add New Source') or normalize-space()='Add New Source']")
	                    ));

	                    js.executeScript("arguments[0].scrollIntoView(true);", addNew);
	                    Thread.sleep(500);
	                    js.executeScript("arguments[0].click();", addNew);

	                    logStep(step++, "Clicked Add New Source", "PASS");

	                } catch (Exception eAdd) {
	                    logStep(step++, "Failed to click Add New Source: " + eAdd.getMessage(), "FAIL");
	                }


	                // ============================
	                // CARD NUMBER
	                // ============================
	                try {
	                    WebElement cardFrame = wait6.until(ExpectedConditions.presenceOfElementLocated(
	                            By.cssSelector("iframe[name*='card'], iframe[id*='card'], iframe[src*='card']")));
	                    driver.switchTo().frame(cardFrame);

	                    WebElement cardNumber = wait6.until(ExpectedConditions.elementToBeClickable(By.name("cardnumber")));
	                    js.executeScript("arguments[0].scrollIntoView(true);", cardNumber);
	                    cardNumber.sendKeys("4111111111111111");

	                    driver.switchTo().defaultContent();
	                    logStep(step++, "Entered Card Number", "PASS");

	                } catch (Exception eCard) {
	                    logStep(step++, "Failed to enter card number: " + eCard.getMessage(), "FAIL");
	                }


	                // ============================
	                // EXPIRY DATE
	                // ============================
	                try {
	                    WebElement expFrame = wait6.until(ExpectedConditions.presenceOfElementLocated(
	                            By.cssSelector("iframe[name*='exp'], iframe[id*='exp'], iframe[title*='exp']")
	                    ));

	                    driver.switchTo().frame(expFrame);

	                    WebElement expiration = wait6.until(ExpectedConditions.elementToBeClickable(
	                            By.cssSelector("input[name='exp-date'], input[autocomplete='cc-exp']")
	                    ));

	                    expiration.sendKeys("0240"); // MMYY
	                    driver.switchTo().defaultContent();

	                    logStep(step++, "Entered Expiry Date", "PASS");

	                } catch (Exception eExp) {
	                    logStep(step++, "Failed to enter expiry: " + eExp.getMessage(), "FAIL");
	                }


	                // ============================
	                // CVC
	                // ============================
	                try {
	                    
	                    String cvcIFrameLocator = "[title*='CVC']"; 

	                    
	                    WebElement cvcFrame = wait6.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(cvcIFrameLocator)));
	                    
	                    
	                    driver.switchTo().frame(cvcFrame);

	                   
	                    WebElement cvcInput = wait6.until(ExpectedConditions.elementToBeClickable(By.name("cvc")));
	                    
	                    
	                    cvcInput.sendKeys("123");

	                    
	                    driver.switchTo().defaultContent();
	                    logStep(step++, "Entered CVC successfully", "PASS");

	                } catch (Exception eCVC) {
	                    
	                    logStep(step++, "Failed to enter CVC (Stripe): " + eCVC.getMessage(), "FAIL");
	                }

	                // ============================
	                // SUBMIT PAYMENT
	                // ============================
	                
	                try {
	                    WebDriverWait wait10 = new WebDriverWait(driver, Duration.ofSeconds(10));

	                    WebElement submit = wait10.until(
	                        ExpectedConditions.elementToBeClickable(
	                            By.xpath("//button[normalize-space()='Submit']")
	                        )
	                    );

	                    ((JavascriptExecutor) driver).executeScript("arguments[0].click();", submit);

	                    logStep(step++, "Clicked on Submit button", "PASS");
	                    System.out.println("Clicked on submit button");

	                    Thread.sleep(500);

	                } catch (Exception e) {
	                    logStep(step++, "Failed to click Submit button: " + e.getMessage(), "FAIL");
	                }

	                
//	                try {
//	                    // 1. बेहतर लोकेटर का उपयोग करें (उदाहरण के लिए: Relative XPath)
//	                    WebElement submitButton = wait6.until(ExpectedConditions.presenceOfElementLocated(
//	                            By.xpath("//button[contains(., 'Submit') or contains(., 'Pay')]") // Submit या Pay text वाले बटन को ढूंढेगा
//	                    ));
//
//	                    // 2. बटन तक स्क्रॉल करें (आपका कोड पहले से ही यह कर रहा है)
//	       //             js.executeScript("arguments[0].scrollIntoView(true);", submitButton);
//	                    
//	                    // 3. बटन पर सीधे JavaScript से क्लिक करें
//	      //              js.executeScript("arguments[0].click();", submitButton); // **यह सबसे विश्वसनीय तरीका है**
//
//	                    logStep(step++, "Clicked Submit using JS", "PASS");
//
//	                } catch (Exception eSubmit) {
//	                    logStep(step++, "Failed to click Submit: " + eSubmit.getMessage(), "FAIL");
//	                }


	                // ============================
	                // CLICK SUCCESS OK BUTTON
	                // ============================
	                try {
	                    WebElement okButton = wait6.until(ExpectedConditions.elementToBeClickable(
	                            By.xpath("//*[@id='react-confirm-alert']/div/div/div/div/button")
	                    ));

	                    js.executeScript("arguments[0].scrollIntoView(true);", okButton);
	                    okButton.click();

	                    logStep(step++, "Clicked OK on Payment Success", "PASS");

	                } catch (Exception eOK) {
	                    logStep(step++, "Failed to click OK after success: " + eOK.getMessage(), "FAIL");
	                }

	            } 
	            // ===== MAIN CATCH =====
	            catch (Exception e) {
	                logStep(step++, "Unexpected Error: " + e.getMessage(), "FAIL");
	            } 
	            // ===== FINALLY =====
	            finally {
	                try {
	                    if (driver != null) {
	                        Thread.sleep(4000);
	                        driver.quit();
	                        logStep(step++, "Browser closed", "PASS");
	                    }
	                } catch (Exception ex) {
	                    ex.printStackTrace();
	                }
	            }

	            System.out.println("✅ Automation complete! Payment flow logged successfully!");

	        } catch (Exception e) {
	            e.printStackTrace();
	        } finally {
	            try {
	                if (driver != null) {
	                    Thread.sleep(3000);
	//                    driver.quit();
	                }
	            } catch (Exception ex) {
	                ex.printStackTrace();
	            }
	        }
	        
	        
	//        finally {
	            try {
	                if (driver != null) {
	                    Thread.sleep(3000);
	                }
	            } catch (Exception ex) {
	                ex.printStackTrace();
	            }

	            // Save Excel File
	            try (FileOutputStream outputStream = new FileOutputStream("AutomationSteps.xlsx")) {
	                workbook.write(outputStream);
	                workbook.close();
	                System.out.println("✅ Excel file saved successfully!");
	            } catch (Exception ex) {
	                ex.printStackTrace();
	            }
	        }
	    }
	

	            
	        
	 
	    



	