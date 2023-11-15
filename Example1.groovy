import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import com.kms.katalon.keyword.excel.ExcelKeywords as ExcelKeywords

WebUI.callTestCase(findTestCase('xxx/xxx - ImpersonalizarLogin'), [:], FailureHandling.STOP_ON_FAILURE)

WebUI.click(findTestObject('Object Repository/Page_JUL/menu'))

WebUI.mouseOver(findTestObject('Page_JUL/relatoriosMenu'))

WebUI.click(findTestObject('Page_JUL/relatoriosDashboard'))

WebUI.click(findTestObject('Page_JUL/navegarTabRelatorios'))

WebUI.verifyElementText(findTestObject('Page_JUL/relatoriosTitulo'), 'Geração de Relatórios')

WebUI.waitForElementPresent(findTestObject('Page_JUL/selecioneUmPortoDM'), 60)

WebUI.click(findTestObject('Page_JUL/selecioneUmPortoDM'))

WebUI.click(findTestObject('Page_JUL/selecioneUmPortoSinesDM'))

WebUI.verifyElementPresent(findTestObject('Page_JUL/relatoriosTabRelatorios'), 0)

WebUI.verifyElementClickable(findTestObject('Page_JUL/relatoriosTabRelatorios'), FailureHandling.STOP_ON_FAILURE)

WebUI.verifyElementPresent(findTestObject('Page_JUL/relatoriosTabConsultas'), 0)

WebUI.verifyElementClickable(findTestObject('Page_JUL/relatoriosTabConsultas'), FailureHandling.STOP_ON_FAILURE)

WebUI.click(findTestObject('Page_JUL/relatorioControloDeFaturacao'))

WebUI.click(findTestObject('Page_JUL/adicionarParametros'))

WebUI.click(findTestObject('Page_JUL/selecioneUmaOpcaoDropdown'))

WebUI.click(findTestObject('Page_JUL/sinesOpcaoDropdown'))

Date todayEDE = new Date() - 30

String todayLess30days = todayEDE.format('dd/MM/yyyy')

WebUI.setText(findTestObject('Page_JUL/relatorioCFEDE'), todayLess30days)

WebUI.sendKeys(findTestObject('Page_JUL/relatorioCFEDE'), Keys.chord(Keys.TAB))

WebUI.sendKeys(findTestObject('Page_JUL/relatorioCFEDEHora'), '1002')

Date todayEATE = new Date() - 20

String todayLess20days = todayEATE.format('dd/MM/yyyy')

WebUI.setText(findTestObject('Page_JUL/relatorioCFEATE'), todayLess20days)

WebUI.sendKeys(findTestObject('Page_JUL/relatorioCFEATE'), Keys.chord(Keys.TAB))

WebUI.sendKeys(findTestObject('Page_JUL/relatorioCFEATEHora'), '1830')

Date todaySDE = new Date() - 10

String todayLess10days = todaySDE.format('dd/MM/yyyy')

WebUI.setText(findTestObject('Page_JUL/relatorioCFSDE'), todayLess10days)

WebUI.sendKeys(findTestObject('Page_JUL/relatorioCFSDE'), Keys.chord(Keys.TAB))

WebUI.sendKeys(findTestObject('Page_JUL/relatorioCFSDEHora'), '0932')

Date todaySATE = new Date()

String today = todaySATE.format('dd/MM/yyyy')

WebUI.setText(findTestObject('Page_JUL/relatorioCFSATE'), today)

WebUI.sendKeys(findTestObject('Page_JUL/relatorioCFSATE'), Keys.chord(Keys.TAB))

WebUI.sendKeys(findTestObject('Page_JUL/relatorioCFSATEHora'), '1700')

WebUI.sendKeys(findTestObject('Page_JUL/numeroEscala'), 'xxx')

WebUI.sendKeys(findTestObject('Page_JUL/relatorioCFIMO'), '123456')

WebUI.sendKeys(findTestObject('Page_JUL/nifDoAgente'), '123456789')

WebUI.click(findTestObject('Page_JUL/relatorioCFSemFaturas'))

WebUI.scrollToElement(findTestObject('Page_JUL/selecioneUmaOpcaoDropdown'), 0)

WebUI.click(findTestObject('Page_JUL/selecioneUmaOpcaoDropdown'))

WebUI.click(findTestObject('Page_JUL/relatorioCFTipoPreFaturaOpcao'))

WebUI.sendKeys(findTestObject('Page_JUL/numeroPreFatura'), 'xxx')

WebUI.click(findTestObject('Page_JUL/selecioneUmaOpcaoDropdown'))

WebUI.scrollToElement(findTestObject('Page_JUL/relatorioCFEstadoPreFatura'), 0)

WebUI.click(findTestObject('Page_JUL/relatorioCFEstadoPreFatura'))

WebUI.sendKeys(findTestObject('Page_JUL/numeroFatura'), 'xxx')

WebUI.click(findTestObject('Page_JUL/fecharSidebar'))

def portoSelecionado = WebUI.getText(findTestObject('Page_JUL/valorSelecionado'))

def eDESelecionado = WebUI.getText(findTestObject('Page_JUL/valor2Selecionado'))

def eDATESelecionado = WebUI.getText(findTestObject('Page_JUL/valor3Selecionado'))

def sDESelecionado = WebUI.getText(findTestObject('Page_JUL/valor4Selecionado'))

def sATESelecionado = WebUI.getText(findTestObject('Page_JUL/valor5Selecionado'))

def numeroEscalaSelecionado = WebUI.getText(findTestObject('Page_JUL/valor6Selecionado'))

def iMOSelecionado = WebUI.getText(findTestObject('Page_JUL/valor7Selecionado'))

def nifDoAgenteSelecionado = WebUI.getText(findTestObject('Page_JUL/valor8Selecionado'))

WebUI.verifyElementVisible(findTestObject('Page_JUL/valorCFSemFaturasSelecionado'))

def tipoPreFaturaSelecionado = WebUI.getText(findTestObject('Page_JUL/valor10Selecionado'))

def numeroPreFaturaSelecionado = WebUI.getText(findTestObject('Page_JUL/valor11Selecionado'))

def estadoPreFaturaSelecionado = WebUI.getText(findTestObject('Page_JUL/valor12Selecionado'))

def numeroFaturaSelecionado = WebUI.getText(findTestObject('Page_JUL/valor13Selecionado'))

WebUI.click(findTestObject('Page_JUL/menuHamburgerRelatorios'))

WebUI.click(findTestObject('Page_JUL/gerarRelatorio'))

WebUI.waitForElementPresent(findTestObject('Page_JUL/relatoriosTabConsultas'), 60)

WebUI.click(findTestObject('Page_JUL/relatoriosTabConsultas'))

WebUI.verifyElementText(findTestObject('Page_JUL/primeiraOpcao'), 'Controlo de Faturação', FailureHandling.STOP_ON_FAILURE)

WebUI.waitForElementPresent(findTestObject('Page_JUL/estadoProcessado'), 180)

WebUI.refresh()

WebUI.click(findTestObject('Page_JUL/relatoriosTabConsultas'))

WebUI.verifyElementVisible(findTestObject('Page_JUL/estadoProcessado'))

WebUI.verifyElementVisible(findTestObject('Page_JUL/fileDownload'))

WebUI.doubleClick(findTestObject('Page_JUL/fileDownload'))

WebUI.click(findTestObject('Page_JUL/ultimaLinhaGerada'))

def nomeRelatorio = WebUI.getText(findTestObject('Page_JUL/relatorioNomeRelatorio'))

def dataProcessamento = WebUI.getText(findTestObject('Page_JUL/relatorioDataDeProcessamento'))

def portoGerado = WebUI.getText(findTestObject('Page_JUL/valorGerado'))

def eDEGerado = WebUI.getText(findTestObject('Page_JUL/valor2Gerado'))

def eDATEGerado = WebUI.getText(findTestObject('Page_JUL/valor3Gerado'))

def sDEGerado = WebUI.getText(findTestObject('Page_JUL/valor4Gerado'))

def sATEGerado = WebUI.getText(findTestObject('Page_JUL/valor5Gerado'))

def numeroEscalaGerado = WebUI.getText(findTestObject('Page_JUL/valor6Gerado'))

def iMOGerado = WebUI.getText(findTestObject('Page_JUL/valor7Gerado'))

def nifDoAgenteGerado = WebUI.getText(findTestObject('Page_JUL/valor8Gerado'))

WebUI.verifyElementVisible(findTestObject('Page_JUL/valorCFSemFaturasGerado'))

def tipoPreFaturaGerado = WebUI.getText(findTestObject('Page_JUL/valor10Gerado'))

def numeroPreFaturaGerado = WebUI.getText(findTestObject('Page_JUL/valor11Gerado'))

def estadoPreFaturaGerado = WebUI.getText(findTestObject('Page_JUL/valor12Gerado'))

def numeroFaturaGerado = WebUI.getText(findTestObject('Page_JUL/valor13Gerado'))

WebUI.verifyMatch(portoSelecionado, portoGerado, false)

WebUI.verifyMatch(eDESelecionado, eDEGerado, false)

WebUI.verifyMatch(eDATESelecionado, eDATEGerado, false)

WebUI.verifyMatch(sDESelecionado, sDEGerado, false)

WebUI.verifyMatch(sATESelecionado, sATEGerado, false)

WebUI.verifyMatch(numeroEscalaSelecionado, numeroEscalaGerado, false)

WebUI.verifyMatch(iMOSelecionado, iMOGerado, false)

WebUI.verifyMatch(nifDoAgenteSelecionado, nifDoAgenteGerado, false)

WebUI.verifyElementVisible(findTestObject('Page_JUL/valorCFSemFaturasGerado'))

WebUI.verifyMatch(tipoPreFaturaSelecionado, tipoPreFaturaGerado, false)

WebUI.verifyMatch(numeroPreFaturaSelecionado, numeroPreFaturaGerado, false)

WebUI.verifyMatch(estadoPreFaturaSelecionado, estadoPreFaturaGerado, false)

WebUI.verifyMatch(numeroFaturaSelecionado, numeroFaturaGerado, false)

WebUI.verifyElementPresent(findTestObject('Page_JUL/fileDetalhes'), 0)

WebUI.verifyElementClickable(findTestObject('Page_JUL/fileDetalhes'), FailureHandling.STOP_ON_FAILURE)

WebUI.scrollToElement(findTestObject('Page_JUL/fileDetalhes'), 0)

def nomeFicheiro = WebUI.getText(findTestObject('Page_JUL/fileDetalhes'))

def SheetDefaultPesquisa = ExcelKeywords.getExcelSheet(('C:\\Users\\xxx\\Downloads\\' + nomeFicheiro) + '.xlsx')

def valorCell1 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 0, 1)

def valorCell2 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 3, 3)

def valorCell3 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 5, 3)

def valorCell4 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 5, 5)

def valorCell5 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 7, 3)

def valorCell6 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 7, 5)

def valorCell7 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 9, 3)

def valorCell8 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 9, 5)

def valorCell9 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 11, 3)

def valorCell10 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 13, 3)

def valorCell11 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 13, 5)

def valorCell12 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 15, 3)

def valorCell13 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 15, 5)

def valorCell14 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 17, 3)

def valorCell15 = ExcelKeywords.getCellByIndex(SheetDefaultPesquisa, 20, 3)

WebUI.verifyMatch(nomeRelatorio, valorCell1.toString(), false)

WebUI.verifyMatch('xxx | ' + portoGerado, valorCell2.toString(), false)

WebUI.verifyMatch(eDEGerado, valorCell3.toString(), false)

WebUI.verifyMatch(eDATEGerado, valorCell4.toString(), false)

WebUI.verifyMatch(sDEGerado, valorCell5.toString(), false)

WebUI.verifyMatch(sATEGerado, valorCell6.toString(), false)

WebUI.verifyMatch(numeroEscalaGerado, valorCell7.toString(), false)

WebUI.verifyMatch(iMOGerado, valorCell8.toString(), false, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.verifyMatch(nifDoAgenteGerado + ' | xxx', valorCell9.toString(), false)

WebUI.verifyMatch('xxx | ' + tipoPreFaturaGerado, valorCell11.toString(), false, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.verifyMatch(numeroPreFaturaGerado, valorCell12.toString(), false, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.verifyMatch('xxx | ' + estadoPreFaturaGerado, valorCell13.toString(), false, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.verifyMatch(numeroFaturaGerado, valorCell14.toString(), false, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.verifyMatch(dataProcessamento, valorCell15.toString(), false, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.closeBrowser()

