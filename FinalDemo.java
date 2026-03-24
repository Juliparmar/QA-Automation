import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.Locale;
import java.util.concurrent.TimeoutException;

import org.apache.commons.io.FileUtils;
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
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class FinalDemo{
	
	
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
    
    
    public static void main(String[] args) throws InterruptedException {

        initializeHeaders();
        WebDriver driver = null;

        try {
            ChromeOptions options = new ChromeOptions();
            options.addArguments("--start-maximized");
            driver = new ChromeDriver(options);

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
            Actions action = new Actions(driver);

            int step = 1;

            // Open URL
            driver.get("https://dev-orders.richervalues.com/auth/login");
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("username")));
            logStep(step++, "Browser opened and URL loaded", "PASS");

            // Username
            WebElement username = wait.until(
                    ExpectedConditions.elementToBeClickable(By.id("username"))
            );
            username.sendKeys("Test_Administrator");
            logStep(step++, "Entered username", "PASS");

            // Password
            WebElement password = wait.until(
                    ExpectedConditions.elementToBeClickable(By.id("password"))
            );
            password.sendKeys("Browse]1Pope|8Vivacious/6Battered+1Swiftness");
            logStep(step++, "Entered password", "PASS");
            Thread.sleep(500);

            // Sign In button
            
            wait.until(ExpectedConditions.invisibilityOfElementLocated(
                    By.xpath("//div[contains(@class,'loader') or contains(@class,'spinner')]")
            ));

            // stable locator
            By signinBtn = By.xpath("//button[@type='submit' and contains(normalize-space(),'Sign In')]");

            WebElement signin = wait.until(
                    ExpectedConditions.presenceOfElementLocated(signinBtn)
            );

            // scroll + JS click
            JavascriptExecutor js = (JavascriptExecutor) driver;
            js.executeScript("arguments[0].scrollIntoView({block:'center'});", signin);
            js.executeScript("arguments[0].click();", signin);
            
            
           
            
            // Place New Order
            WebElement placeNewOrder = wait.until(
                    ExpectedConditions.elementToBeClickable(By.linkText("Place New Order"))
            );
            action.moveToElement(placeNewOrder).click().perform();
            logStep(step++, "Clicked on 'Place New Order'", "PASS");

            // Handle "Got it" popup (if present)
            try {
                WebElement gotItButton = wait.until(
                        ExpectedConditions.elementToBeClickable(
                                By.xpath("//button[contains(text(),'Got it')] | //button[contains(text(),'Got it, thank you')]")
                        )
                );
                gotItButton.click();
                logStep(step++, "Clicked 'Got it, thank you' popup", "PASS");

            } catch (Exception e) {
                logStep(step++, "'Got it' popup not displayed", "INFO");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

     // Correct Actions initialization
     Actions action = new Actions(driver);

     // Step counter
     int step = 1;

        
        
        // Evaluations menu
       
		
		LocalDate wait2;
		WebElement evaluations = wait.until(
		        ExpectedConditions.elementToBeClickable(By.id("eval"))
		);
    //    Actions action2;
		action.moveToElement(evaluations).perform();
        evaluations.click();
   //     int step;
		logStep(step++, "Clicked on 'Evaluations'", "PASS");
        

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
        WebElement loanNumber = wait.until(ExpectedConditions.elementToBeClickable(By.id("client_loan_number")));
        loanNumber.sendKeys("LN-00098765");
        logStep(step++, "Entered Loan Number: LN-00098765", "PASS");
 //       Thread.sleep(2000);
        System.out.println(); 
        
        // Property selection
        WebDriverWait wait4 = new WebDriverWait(driver, Duration.ofSeconds(30));

        By vacantInput = By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/div[2]/form/div[1]/div[3]/div[1]/div/div[2]/label");

        WebElement propertyVacant = wait.until(
                ExpectedConditions.presenceOfElementLocated(vacantInput)
        );

        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("arguments[0].scrollIntoView({block:'center'});", propertyVacant);
        js.executeScript("arguments[0].click();", propertyVacant);

        logStep(step++, "Selected 'Property Vacant Land'", "PASS");
//        Thread.sleep(500);

        WebElement propertyUndergoing = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.xpath("//*[@id='root']/div[1]/div[2]/div/div[2]/form/div[1]/div[3]/div[2]/div/div[2]/label")
                )
        );

        propertyUndergoing.click();
        logStep(step++, "Selected 'Property Currently Undergoing'", "PASS");
        
        // Report Type
        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(40));

        By reportTypeLocator =
                By.xpath("//label[contains(normalize-space(),'Reno')]");

        WebElement reportType = wait1.until(
                ExpectedConditions.visibilityOfElementLocated(reportTypeLocator)
        );

//        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("arguments[0].scrollIntoView({block:'center'});", reportType);
        js.executeScript("arguments[0].click();", reportType);

        logStep(step++, "Selected 'Reno ARV' report type", "PASS");
     

      
//         Inspection Needed
//        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("window.scrollBy(0, 500)");
        logStep(step++, "Page scrolled down 500px", "PASS");
        Thread.sleep(3000);
        

        // Inspection Needed
 //       JavascriptExecutor js = (JavascriptExecutor) driver;
//        js.executeScript("window.scrollBy(0, 500)");
//        logStep(step++, "Page scrolled down 500px", "PASS");
//        Thread.sleep(3000);
//        

//        JavascriptExecutor js = (JavascriptExecutor) driver;
//         js.executeScript("window.scrollBy(0, 500)");
//         Thread.sleep(1000);
        
        WebDriverWait wait3 = new WebDriverWait(driver, Duration.ofSeconds(20));

     // Wait for loader to disappear
     wait3.until(ExpectedConditions.invisibilityOfElementLocated(By.id("loader")));

     // Now wait for your element to be clickable
     WebElement inspectionNeeded = wait.until(
         ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"inspection_type_0\"]/div/label"))
     );

     // Try clicking
     try {
         inspectionNeeded.click();
         System.out.println("Clicked on 'Inspection Needed' successfully");
     } catch (Exception e) {
         // Fallback if normal click fails
//         JavascriptExecutor js = (JavascriptExecutor) driver;
         js.executeScript("arguments[0].click();", inspectionNeeded);
         System.out.println("Clicked using JS Executor as fallback");
     }
    
        

        // Turnaround Time
        WebDriverWait wait25 = new WebDriverWait(driver, Duration.ofSeconds(10));
        WebElement TurnaroundTimeLabel = wait25.until(
            ExpectedConditions.elementToBeClickable(By.xpath("//label[@for='radio-turnaround_time-standard']"))
        );
        TurnaroundTimeLabel.click();
        logStep(step++, "Selected standard Turnaround Time", "PASS");
//        Thread.sleep(3000);

        
        // Closing Date
//        WebElement ClosingDate = driver.findElement(By.name("select_date"));
//        ClosingDate.click();
//        logStep(step++, "Clicked on Closing Date calendar input", "PASS");
//        Thread.sleep(1000);


        //
        try {
            WebElement calendarInput = wait1.until(
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

            WebElement dayEl = wait1.until(ExpectedConditions.elementToBeClickable(
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
            
            
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(".rbt-menu")));
            selectClient.sendKeys(Keys.ARROW_DOWN);
            selectClient.sendKeys(Keys.ENTER);
            logStep(step++, "Test_2 Test_2", "PASS");
            Thread.sleep(1000);

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
             
          
             
//             // Effective Date
//             
//             WebElement label = wait.until(
//            	        ExpectedConditions.elementToBeClickable(
//            	                By.xpath("//label[@for='radio-turnaround_time-standard']")
//            	        )
//            	);
//            	js.executeScript("arguments[0].click();", label);
//             

        // Lock box
             
             WebDriverWait wait8 = new WebDriverWait(driver, Duration.ofSeconds(15));
    //         JavascriptExecutor js = (JavascriptExecutor) driver;

             WebElement element = wait8.until(
                     ExpectedConditions.elementToBeClickable(
                             By.xpath("//*[@id='root']/div[1]/div[2]/div/div[2]/form/div[1]/div[11]/div/div/fieldset/div[1]/div[2]/label")
                     )
             );

             // Scroll into view
             js.executeScript("arguments[0].scrollIntoView({block:'center'});", element);

             // Click using JS (label / radio safe)
             js.executeScript("arguments[0].click();", element);

             logStep(step++, "Clicked element using wait", "PASS");
             
             
             
             WebDriverWait wait45 = new WebDriverWait(driver, Duration.ofSeconds(20));

             // Wait for loader to disappear
             wait45.until(ExpectedConditions.invisibilityOfElementLocated(By.id("loader")));

             WebElement contact = wait.until(
            	        ExpectedConditions.visibilityOfElementLocated(
            	                By.id("name_0")
            	        )
            	);

            	contact.click();
    //        	contact.clear();
            	contact.sendKeys("John");

            	System.out.println("Contact name entered successfully");
            	
            	
            	WebElement email = wait.until(
            	        ExpectedConditions.visibilityOfElementLocated(
            	                By.id("email_0")
            	        )
            	);

            	email.click();
      //      	email.clear();
            	email.sendKeys("John@email.com");

            	System.out.println("Email entered successfully");
            	
	           WebElement contPhone = wait.until(
                ExpectedConditions.visibilityOfElementLocated(
                        By.id("phone_0")
                )
        );

        contPhone.click();
  //      contPhone.clear();
        contPhone.sendKeys("2536256325");

        System.out.println("Contact phone entered successfully");
        
        
//	           
        // Lender Entity
             WebElement lenderDropdown = wait1.until(ExpectedConditions.elementToBeClickable(By.id("company_entity_id")));
             logStep(step++, "Opened Lender Entity dropdown", "PASS");

             // Use Select class to choose option by visible text
             Select lenderSelect = new Select(lenderDropdown);
             lenderSelect.selectByVisibleText("Anhas Client Company -Dev");
             logStep(step++, "Selected Lender Entity option", "PASS");
      
  
        // Scroll again
        js.executeScript("window.scrollBy(0, 500)");
        logStep(step++, "Scrolled down 500px before Next", "PASS");
 //       Thread.sleep(1000);


        // Next Button
        WebDriverWait wait11 = new WebDriverWait(driver, Duration.ofSeconds(10));
        WebElement nextBtn = wait11.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='Next']/parent::button")));
        nextBtn.click();
        
        
      
        
        
//      Step 2
       
        WebDriverWait wait46 = new WebDriverWait(driver, Duration.ofSeconds(20));

        // Wait for loader to disappear
        wait46.until(ExpectedConditions.invisibilityOfElementLocated(By.id("loader")));

     // Property Type
        /* =========================
        Property Type - Single
        ========================= */
     WebElement propertyType = wait.until(
             ExpectedConditions.elementToBeClickable(
                     By.xpath("//label[contains(text(),'Single')]")
             )
     );
     js.executeScript("arguments[0].scrollIntoView({block:'center'});", propertyType);
     js.executeScript("arguments[0].click();", propertyType);
     logStep(step++, "Selected property type: Single", "PASS");


     /* =========================
        Address Input
        ========================= */
     WebElement addressInput = wait.until(
             ExpectedConditions.visibilityOfElementLocated(
                     By.className("pac-target-input")
             )
     );
     addressInput.clear();
     addressInput.sendKeys("2560 DUKELAND DR");
     logStep(step++, "Entered address: 2560 DUKELAND DR", "PASS");


     /* =========================
        First Address Suggestion
        ========================= */
     WebElement firstSuggestion = wait.until(
             ExpectedConditions.visibilityOfElementLocated(
                     By.className("pac-item")
             )
     );
     firstSuggestion.click();
     logStep(step++, "Selected first suggested address", "PASS");


     /* =========================
        Property Subtype - Single Family
        ========================= */
     WebElement proType = wait.until(
             ExpectedConditions.elementToBeClickable(
                     By.xpath("//label[contains(text(),'Single-Family')]")
             )
     );
     js.executeScript("arguments[0].click();", proType);
     logStep(step++, "Selected property subtype: Single-Family", "PASS");


     /* =========================
        Condition - Moderate
        ========================= */
//     WebElement condition = wait.until(
//             ExpectedConditions.elementToBeClickable(
//                     By.xpath("//td[contains(text(),'Moderate')]")
//             )
//     );
//     condition.click();
//     logStep(step++, "Selected property condition: Moderate", "PASS");


     js.executeScript("window.scrollBy(0, 500)");
     logStep(step++, "Page scrolled down 500px", "PASS");
     /* =========================
        Numeric Fields
        ========================= */
     wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("above_grade_sqft"))).sendKeys("3000");
     wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("bedrooms"))).sendKeys("3");
     wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("bathrooms"))).sendKeys("2");
     wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("year_built"))).sendKeys("1994");
     wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("stories"))).sendKeys("");
     wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("lot_size_square_feet"))).sendKeys("20000");
     wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("garage_spaces"))).sendKeys("1");

     logStep(step++, "Entered property numeric details", "PASS");


     /* =========================
        Renovation Budget
        ========================= */
     WebElement budget = wait.until(
             ExpectedConditions.visibilityOfElementLocated(
                     By.id("borrower_budget")
             )
     );
     budget.clear();
     budget.sendKeys("2500000");
     logStep(step++, "Entered Renovation Budget", "PASS");


     /* =========================
        Valuation Report Comment
        ========================= */
     WebElement valuationReport = wait.until(
             ExpectedConditions.visibilityOfElementLocated(
                     By.id("valuation_commentary_or_instruction")
             )
     );
     valuationReport.sendKeys("Additional comment");
     logStep(step++, "Entered Valuation Report comment", "PASS");

        // Next Button
        WebDriverWait wait1111 = new WebDriverWait(driver, Duration.ofSeconds(10));
        WebElement Nextbtn = wait1111.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[contains(.,'Next')]")));
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

        	    
//        	    Actions action = new Actions(driver);
        	    action.moveToElement(fixLaterButton).perform();
        	    System.out.println("✅ Mouse moved to 'Fix Later' button");

        	    // Click the button
        	    fixLaterButton.click();
        	    System.out.println("✅ 'Fix Later' button clicked successfully");

        	    Thread.sleep(500); 

        	} catch (Exception e) {
        	    System.out.println("⚠️ Could not click 'Fix Later' button: " + e.getMessage());
        	}
        
        
//	       Step 3
         
        WebElement Nextbtn2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/form/div[2]/div/div[3]/button[2]"));
        Nextbtn2.click();
        logStep(step++, "Clicked on Next button (Final Step of this page)", "PASS");
 //       Thread.sleep(500);
        
//       Step 4
         
        js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
        logStep(step++, "Scrolled to bottom of the page", "PASS");

        WebElement save = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/div[2]/div/form/div[2]/div/div[3]/button"));
        save.click();
        logStep(step++, "Clicked on Save button", "PASS");

  //      Thread.sleep(500);
    

         
//       Step 5
         
         
        WebDriverWait wait47 = new WebDriverWait(driver, Duration.ofSeconds(10));

        // Wait for loader to disappear
        wait47.until(ExpectedConditions.invisibilityOfElementLocated(By.id("loader")));


            // ============================
            // PAYMENT BUTTON
            // ============================
            WebDriverWait wait5 = new WebDriverWait(driver, Duration.ofSeconds(10));

            WebElement paymentBtn = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.xpath("//button[normalize-space()='Select Payment Option']")
                    )
            );
            js.executeScript("arguments[0].scrollIntoView(true);", paymentBtn);
            Thread.sleep(2000);
            js.executeScript("arguments[0].click();", paymentBtn);
            logStep(step++, "Clicked Payment button", "PASS");

        
            // ============================
            // SELECT CREDIT CARD OPTION
            // ============================
            WebDriverWait wait6 = new WebDriverWait(driver, Duration.ofSeconds(10));

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

            
//            try {
//                // 1. बेहतर लोकेटर का उपयोग करें (उदाहरण के लिए: Relative XPath)
//                WebElement submitButton = wait6.until(ExpectedConditions.presenceOfElementLocated(
//                        By.xpath("//button[contains(., 'Submit') or contains(., 'Pay')]") // Submit या Pay text वाले बटन को ढूंढेगा
//                ));
//
//                // 2. बटन तक स्क्रॉल करें (आपका कोड पहले से ही यह कर रहा है)
//   //             js.executeScript("arguments[0].scrollIntoView(true);", submitButton);
//                
//                // 3. बटन पर सीधे JavaScript से क्लिक करें
//  //              js.executeScript("arguments[0].click();", submitButton); // **यह सबसे विश्वसनीय तरीका है**
//
//                logStep(step++, "Clicked Submit using JS", "PASS");
//
//            } catch (Exception eSubmit) {
//                logStep(step++, "Failed to click Submit: " + eSubmit.getMessage(), "FAIL");
//            }


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

         
        
            try {
                // ===== MAIN TEST STEPS =====
                System.out.println("Automation running...");

                // 👉 yahan aapke actual automation steps honge

            } catch (Exception e) {
                logStep(step++, "Unexpected Error: " + e.getMessage(), "FAIL");
                e.printStackTrace();

            } 

                // ===== SAVE EXCEL FIRST =====
                try (FileOutputStream outputStream =
                             new FileOutputStream("AutomationSteps.xlsx")) {

                    workbook.write(outputStream);
                    workbook.close();
                    System.out.println("✅ Excel file saved successfully!");

                } catch (Exception ex) {
                    ex.printStackTrace();
                }

                // ===== CLOSE BROWSER =====
                try {
                    if (driver != null) {
                        Thread.sleep(3000);
                        driver.quit();
                        logStep(step++, "Browser closed", "PASS");
                    }
                } catch (Exception ex) {
                    ex.printStackTrace();  
                }
            

            // ===== FINAL MESSAGE =====
            System.out.println("✅ Automation complete! Payment flow logged successfully!");
    }
}
   
