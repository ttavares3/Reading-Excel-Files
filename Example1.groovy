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

WebUI.click(findTestObject('Object Repository/Page_xxx/menu'))

WebUI.mouseOver(findTestObject('Page_xxx/relatoriosMenu'))

WebUI.click(findTestObject('Page_xxx/relatoriosDashboard'))

WebUI.click(findTestObject('Page_xxx/navegarTabRelatorios'))

WebUI.verifyElementText(findTestObject('Page_xxx/relatoriosTitulo'), 'Geração de Relatórios')

WebUI.waitForElementPresent(findTestObject('Page_xxx/selecioneUmPortoDM'), 60)

WebUI.click(findTestObject('Page_xxx/selecioneUmPortoDM'))

WebUI.click(findTestObject('Page_xxx/selecioneUmPortoSinesDM'))

WebUI.verifyElementPresent(findTestObject('Page_xxx/relatoriosTabRelatorios'), 0)

WebUI.verifyElementClickable(findTestObject('Page_xxx/relatoriosTabRelatorios'), FailureHandling.STOP_ON_FAILURE)

WebUI.verifyElementPresent(findTestObject('Page_xxx/relatoriosTabConsultas'), 0)

WebUI.verifyElementClickable(findTestObject('Page_xxx/relatoriosTabConsultas'), FailureHandling.STOP_ON_FAILURE)

WebUI.click(findTestObject('Page_xxx/relatorioControloDeFaturacao'))

WebUI.click(findTestObject('Page_xxx/adicionarParametros'))

WebUI.click(findTestObject('Page_xxx/selecioneUmaOpcaoDropdown'))

WebUI.click(findTestObject('Page_xxx/sinesOpcaoDropdown'))

Date todayEDE = new Date() - 30

String todayLess30days = todayEDE.format('dd/MM/yyyy')

WebUI.setText(findTestObject('Page_xxx/relatorioCFEDE'), todayLess30days)

WebUI.sendKeys(findTestObject('Page_xxx/relatorioCFEDE'), Keys.chord(Keys.TAB))

WebUI.sendKeys(findTestObject('Page_xxx/relatorioCFEDEHora'), '1002')

Date todayEATE = new Date() - 20

String todayLess20days = todayEATE.format('dd/MM/yyyy')

WebUI.setText(findTestObject('Page_xxx/relatorioCFEATE'), todayLess20days)

WebUI.sendKeys(findTestObject('Page_xxx/relatorioCFEATE'), Keys.chord(Keys.TAB))

WebUI.sendKeys(findTestObject('Page_xxx/relatorioCFEATEHora'), '1830')

Date todaySDE = new Date() - 10

String todayLess10days = todaySDE.format('dd/MM/yyyy')

WebUI.setText(findTestObject('Page_xxx/relatorioCFSDE'), todayLess10days)

WebUI.sendKeys(findTestObject('Page_xxx/relatorioCFSDE'), Keys.chord(Keys.TAB))

WebUI.sendKeys(findTestObject('Page_xxx/relatorioCFSDEHora'), '0932')

Date todaySATE = new Date()

String today = todaySATE.format('dd/MM/yyyy')

WebUI.setText(findTestObject('Page_xxx/relatorioCFSATE'), today)

WebUI.sendKeys(findTestObject('Page_xxx/relatorioCFSATE'), Keys.chord(Keys.TAB))

WebUI.sendKeys(findTestObject('Page_xxx/relatorioCFSATEHora'), '1700')

WebUI.sendKeys(findTestObject('Page_xxx/numeroEscala'), 'xxx')

WebUI.sendKeys(findTestObject('Page_xxx/relatorioCFIMO'), '123456')

WebUI.sendKeys(findTestObject('Page_xxx/nifDoAgente'), '123456789')

WebUI.click(findTestObject('Page_xxx/relatorioCFSemFaturas'))

WebUI.scrollToElement(findTestObject('Page_xxx/selecioneUmaOpcaoDropdown'), 0)

WebUI.click(findTestObject('Page_xxx/selecioneUmaOpcaoDropdown'))

WebUI.click(findTestObject('Page_xxx/relatorioCFTipoPreFaturaOpcao'))

WebUI.sendKeys(findTestObject('Page_xxx/numeroPreFatura'), 'xxx')

WebUI.click(findTestObject('Page_xxx/selecioneUmaOpcaoDropdown'))

WebUI.scrollToElement(findTestObject('Page_xxx/relatorioCFEstadoPreFatura'), 0)

WebUI.click(findTestObject('Page_xxx/relatorioCFEstadoPreFatura'))

WebUI.sendKeys(findTestObject('Page_xxx/numeroFatura'), 'xxx')

WebUI.click(findTestObject('Page_xxx/fecharSidebar'))

def portoSelecionado = WebUI.getText(findTestObject('Page_xxx/valorSelecionado'))

def eDESelecionado = WebUI.getText(findTestObject('Page_xxx/valor2Selecionado'))

def eDATESelecionado = WebUI.getText(findTestObject('Page_xxx/valor3Selecionado'))

def sDESelecionado = WebUI.getText(findTestObject('Page_xxx/valor4Selecionado'))

def sATESelecionado = WebUI.getText(findTestObject('Page_xxx/valor5Selecionado'))

def numeroEscalaSelecionado = WebUI.getText(findTestObject('Page_xxx/valor6Selecionado'))

def iMOSelecionado = WebUI.getText(findTestObject('Page_xxx/valor7Selecionado'))

def nifDoAgenteSelecionado = WebUI.getText(findTestObject('Page_xxx/valor8Selecionado'))

WebUI.verifyElementVisible(findTestObject('Page_xxx/valorCFSemFaturasSelecionado'))

def tipoPreFaturaSelecionado = WebUI.getText(findTestObject('Page_xxx/valor10Selecionado'))

def numeroPreFaturaSelecionado = WebUI.getText(findTestObject('Page_xxx/valor11Selecionado'))

def estadoPreFaturaSelecionado = WebUI.getText(findTestObject('Page_xxx/valor12Selecionado'))

def numeroFaturaSelecionado = WebUI.getText(findTestObject('Page_xxx/valor13Selecionado'))

WebUI.click(findTestObject('Page_xxx/menuHamburgerRelatorios'))

WebUI.click(findTestObject('Page_xxx/gerarRelatorio'))

WebUI.waitForElementPresent(findTestObject('Page_xxx/relatoriosTabConsultas'), 60)

WebUI.click(findTestObject('Page_xxx/relatoriosTabConsultas'))

WebUI.verifyElementText(findTestObject('Page_xxx/primeiraOpcao'), 'Controlo de Faturação', FailureHandling.STOP_ON_FAILURE)

WebUI.waitForElementPresent(findTestObject('Page_xxx/estadoProcessado'), 180)

WebUI.refresh()

WebUI.click(findTestObject('Page_xxx/relatoriosTabConsultas'))

WebUI.verifyElementVisible(findTestObject('Page_xxx/estadoProcessado'))

WebUI.verifyElementVisible(findTestObject('Page_xxx/fileDownload'))

WebUI.doubleClick(findTestObject('Page_xxx/fileDownload'))

WebUI.click(findTestObject('Page_xxx/ultimaLinhaGerada'))

def nomeRelatorio = WebUI.getText(findTestObject('Page_xxx/relatorioNomeRelatorio'))

def dataProcessamento = WebUI.getText(findTestObject('Page_xxx/relatorioDataDeProcessamento'))

def portoGerado = WebUI.getText(findTestObject('Page_xxx/valorGerado'))

def eDEGerado = WebUI.getText(findTestObject('Page_xxx/valor2Gerado'))

def eDATEGerado = WebUI.getText(findTestObject('Page_xxx/valor3Gerado'))

def sDEGerado = WebUI.getText(findTestObject('Page_xxx/valor4Gerado'))

def sATEGerado = WebUI.getText(findTestObject('Page_xxx/valor5Gerado'))

def numeroEscalaGerado = WebUI.getText(findTestObject('Page_xxx/valor6Gerado'))

def iMOGerado = WebUI.getText(findTestObject('Page_xxx/valor7Gerado'))

def nifDoAgenteGerado = WebUI.getText(findTestObject('Page_xxx/valor8Gerado'))

WebUI.verifyElementVisible(findTestObject('Page_xxx/valorCFSemFaturasGerado'))

def tipoPreFaturaGerado = WebUI.getText(findTestObject('Page_xxx/valor10Gerado'))

def numeroPreFaturaGerado = WebUI.getText(findTestObject('Page_xxx/valor11Gerado'))

def estadoPreFaturaGerado = WebUI.getText(findTestObject('Page_xxx/valor12Gerado'))

def numeroFaturaGerado = WebUI.getText(findTestObject('Page_xxx/valor13Gerado'))

WebUI.verifyMatch(portoSelecionado, portoGerado, false)

WebUI.verifyMatch(eDESelecionado, eDEGerado, false)

WebUI.verifyMatch(eDATESelecionado, eDATEGerado, false)

WebUI.verifyMatch(sDESelecionado, sDEGerado, false)

WebUI.verifyMatch(sATESelecionado, sATEGerado, false)

WebUI.verifyMatch(numeroEscalaSelecionado, numeroEscalaGerado, false)

WebUI.verifyMatch(iMOSelecionado, iMOGerado, false)

WebUI.verifyMatch(nifDoAgenteSelecionado, nifDoAgenteGerado, false)

WebUI.verifyElementVisible(findTestObject('Page_xxx/valorCFSemFaturasGerado'))

WebUI.verifyMatch(tipoPreFaturaSelecionado, tipoPreFaturaGerado, false)

WebUI.verifyMatch(numeroPreFaturaSelecionado, numeroPreFaturaGerado, false)

WebUI.verifyMatch(estadoPreFaturaSelecionado, estadoPreFaturaGerado, false)

WebUI.verifyMatch(numeroFaturaSelecionado, numeroFaturaGerado, false)

WebUI.verifyElementPresent(findTestObject('Page_xxx/fileDetalhes'), 0)

WebUI.verifyElementClickable(findTestObject('Page_xxx/fileDetalhes'), FailureHandling.STOP_ON_FAILURE)

WebUI.scrollToElement(findTestObject('Page_xxx/fileDetalhes'), 0)

def nomeFicheiro = WebUI.getText(findTestObject('Page_xxx/fileDetalhes'))

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

