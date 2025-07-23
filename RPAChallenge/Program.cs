using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

Console.WriteLine("Initializing RPA Challenge...");

// define o caminho de download do arquivo Excel
string downloadPath = Directory.GetCurrentDirectory();
string filePath = downloadPath + "/challenge.xlsx";

// define as configurações do ChromeDriver
var chromeOptions = new ChromeOptions();
chromeOptions.AddArgument("--start-maximized");
chromeOptions.AddUserProfilePreference("download.default_directory", downloadPath);
chromeOptions.AddUserProfilePreference("download.prompt_for_download", false);
chromeOptions.AddUserProfilePreference("download.directory_upgrade", true);
chromeOptions.AddUserProfilePreference("safebrowsing.enabled", true);

// cria uma instância do ChromeDriver
IWebDriver driver = new ChromeDriver(chromeOptions);

// cria uma instância do WebDriverWait para aguardar elementos
WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

// acessa o desafio
driver.Navigate().GoToUrl("https://rpachallenge.com");

Console.WriteLine("Verifying if the workbook is already downloaded...");

// verifica se o arquivo já foi baixado, se não, clica no link para baixar
if (!File.Exists(filePath))
{
    Console.WriteLine("File not found. Downloading the Excel file...");
    var downloadExcel = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[contains(text(),'Download Excel')]")));
    downloadExcel.Click();
    // aguarda o download do arquivo
    wait.Until(driver => File.Exists(filePath));
    Console.WriteLine("File downloaded successfully.");
}
else
{
    Console.WriteLine("File already exists. Skipping download.");
}

// cria uma instância do ClosedXML para manipular o arquivo Excel
var workbook = new XLWorkbook(filePath);
var worksheet = workbook.Worksheet(1);

Console.WriteLine("Starting the challenge...");

// mapeia o botão de início do desafio e clica nele
var startButton = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(text(),'Start')]")));
startButton.Click();

Console.WriteLine("Filling the form with data from the Excel file...");

// lê os dados do Excel, começando da segunda linha (pulando o cabeçalho)
var rows = worksheet.RowsUsed();
foreach(var row in rows.Skip(1).Where(row => !row.IsEmpty())) // pula o cabeçalho e linhas vazias
{
    Console.WriteLine($"Filling form number {row.RowNumber() - 1}...");

    // mapeia os campos do formulário e preenche com os dados
    var firstName = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@ng-reflect-name='labelFirstName']")));
    firstName.SendKeys(row.Cell("A").Value.ToString());
    var lastName = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@ng-reflect-name='labelLastName']")));
    lastName.SendKeys(row.Cell("B").Value.ToString());
    var companyName = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@ng-reflect-name='labelCompanyName']")));
    companyName.SendKeys(row.Cell("C").Value.ToString());
    var roleCompany = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@ng-reflect-name='labelRole']")));
    roleCompany.SendKeys(row.Cell("D").Value.ToString());
    var address = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@ng-reflect-name='labelAddress']")));
    address.SendKeys(row.Cell("E").Value.ToString());
    var email = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@ng-reflect-name='labelEmail']")));
    email.SendKeys(row.Cell("F").Value.ToString());
    var phoneNumber = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@ng-reflect-name='labelPhone']")));
    phoneNumber.SendKeys(row.Cell("G").Value.ToString());
    // mapeia o botão de envio do formulário e clica nele
    var submitButton = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@value='Submit']")));
    submitButton.Click();
}

Console.WriteLine("Form filled and submitted successfully, congratulations!");

var successMessage = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@class='message2']")));
Console.WriteLine(successMessage.Text);

// fecha o navegador
Thread.Sleep(15000); // aguarda 15 segundos para visualizar a mensagem de sucesso
driver.Quit();

// finaliza o programa
Console.WriteLine("Challenge completed. Press any key to exit...");
var exit = Console.ReadKey();