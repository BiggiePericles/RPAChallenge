using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

// define as opções do Chrome
var options = new ChromeOptions();
options.AddArgument("--start-maximized");

// cria uma instância do ChromeDriver
IWebDriver driver = new ChromeDriver(options);

// cria uma instância do WebDriverWait para aguardar elementos
WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));

// acessa o desafio
driver.Navigate().GoToUrl("https://rpachallenge.com");

// mapeia o elemento de download do Excel
var downloadExcel = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[contains(text(),'Download Excel')]")));

// verifica se o arquivo já foi baixado, se não, clica no link para baixar
string filePath = @"C:\Users\leolo\Downloads\challenge.xlsx";
if (!File.Exists(filePath))
{
    downloadExcel.Click();
}

// cria uma instância do ClosedXML para manipular o arquivo Excel
var workbook = new XLWorkbook(filePath);
var worksheet = workbook.Worksheet(1);

// mapeia o botão de início do desafio e clica nele
var startButton = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(text(),'Start')]")));
startButton.Click();

// lê os dados do Excel, começando da segunda linha (pulando o cabeçalho)
var rows = worksheet.RowsUsed();
foreach(var row in rows.Skip(1).Where(row => !row.IsEmpty())) // pula o cabeçalho e linhas vazias
{
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

// aguarda 10 segundos para visualizar o resultado antes de fechar o navegador
Thread.Sleep(10000);
driver.Quit();