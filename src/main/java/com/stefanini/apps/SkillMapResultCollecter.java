package com.stefanini.apps;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAreaChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAreaSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScaling;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.STAxPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLegendPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STOrientation;
import org.openxmlformats.schemas.drawingml.x2006.chart.STTickLblPos;
import org.w3c.dom.css.RGBColor;

public class SkillMapResultCollecter {
	private static final int ROLE_CELL_INDEX = 2;
	private static final int UML_CELL_INDEX = 60;
	private static final int MODELODEDADOS_CELL_INDEX = 59;
	private static final int YAGNI_CELL_INDEX = 58;
	private static final int SOLID_CELL_INDEX = 57;
	private static final int KISS_CELL_INDEX = 56;
	private static final int DRY_CELL_INDEX = 55;
	private static final int DEMANDAS_CELL_INDEX = 54;
	private static final int SIE_CELL_INDEX = 53;
	private static final int SETTLEMENT_CELL_INDEX = 52;
	private static final int RELATORIO_CELL_INDEX = 51;
	private static final int REGULATORIO_CELL_INDEX = 50;
	private static final int PRECO_CELL_INDEX = 49;
	private static final int MONITORACAO_CELL_INDEX = 48;
	private static final int MIGRACAO_CELL_INDEX = 47;
	private static final int FRAUDE_CELL_INDEX = 46;
	private static final int FIC_CELL_INDEX = 45;
	private static final int CONVIVENCIA_CELL_INDEX = 44;
	private static final int CHB_CELL_INDEX = 43;
	private static final int CADASTRO_CELL_INDEX = 42;
	private static final int RELEASEBANDEIRAS_CELL_INDEX = 41;
	private static final int MODELO_NEGOCIO_CELL_INDEX = 40;
	private static final int EMPATIA_CELL_INDEX = 39;
	private static final int AUTOCONHECIMENTO_CELL_INDEX = 38;
	private static final int NEGOCIACAO_CELL_INDEX = 37;
	private static final int LIDERANCA_CELL_INDEX = 36;
	private static final int CMMI_CELL_INDEX = 35;
	private static final int PMBOK_CELL_INDEX = 34;
	private static final int SAFE_CELL_INDEX = 33;
	private static final int XP_CELL_INDEX = 32;
	private static final int MANIFESTO_CELL_INDEX = 31;
	private static final int SCRUM_CELL_INDEX = 30;
	private static final int ANGULAR_CELL_INDEX = 29;
	private static final int HTML5_CELL_INDEX = 28;
	private static final int SVN_CELL_INDEX = 27;
	private static final int GIT_CELL_INDEX = 26;
	private static final int MAVEN_CELL_INDEX = 25;
	private static final int JENKINS_CELL_INDEX = 24;
	private static final int DOCKER_CELL_INDEX = 23;
	private static final int SPARK_CELL_INDEX = 22;
	private static final int BOOT_CELL_INDEX = 21;
	private static final int KAFKA_CELL_INDEX = 20;
	private static final int ZUUL_CELL_INDEX = 19;
	private static final int FEIGN_CELL_INDEX = 18;
	private static final int RIBBON_CELL_INDEX = 17;
	private static final int EUREKA_CELL_INDEX = 16;
	private static final int REST_CELL_INDEX = 15;
	private static final int SOAP_CELL_INDEX = 14;
	private static final int MICROSERVICE_CELL_INDEX = 13;
	private static final int JMS_CELL_INDEX = 12;
	private static final int ORACLE_CELL_INDEX = 11;
	private static final int SQL_CELL_INDEX = 10;
	private static final int JPA_CELL_INDEX = 9;
	private static final int BATCH_CELL_INDEX = 8;
	private static final int SPRING_CELL_INDEX = 7;
	private static final int EJB_CELL_INDEX = 6;
	private static final int JSF_CELL_INDEX = 5;
	private static final int JAVA_CELL_INDEX = 4;
	private static final int NAME_CELL_INDEX = 1;
	private static final boolean BY_ROLE = true;
	
	private static final String SKILL_FILE = "C:\\Stefanini\\PJ00212 - BoB\\Stakeholders\\SkillMap.xlsx";
	private static final String OUTPUT_FOLDER = "C:\\Stefanini\\PJ00212 - BoB\\Stakeholders\\Resultados SkillMap\\";

	@SuppressWarnings({ "resource" })
	public void collect(String skillMapFilePath, String outputFolder) throws Exception {
		File skillMapFile = new File(skillMapFilePath == null ? SKILL_FILE : skillMapFilePath);
		FileInputStream skillMapStream = new FileInputStream(skillMapFile);
		Workbook workbook = new XSSFWorkbook(skillMapStream);
		Sheet sh = workbook.getSheet("skillmap");

		for (int i = 1; i < sh.getPhysicalNumberOfRows(); i++) {
			String name = getCellValue(sh, i, NAME_CELL_INDEX);
			System.out.println("Iniciando processamento de resultados do(a) " + name);
			if (name == "#N/D"
			// || !name.toLowerCase().equals("Filipe Pacheco
			// Souza".toLowerCase())
			)
				continue;

			Integer javaTotal = 0;
			Integer jsfTotal = 0;
			Integer ejbTotal = 0;
			Integer springTotal = 0;
			Integer batchTotal = 0;
			Integer jpaTotal = 0;
			Integer sqlTotal = 0;
			Integer oracleTotal = 0;
			Integer jmsTotal = 0;
			Integer microserviceTotal = 0;
			Integer soapTotal = 0;
			Integer restTotal = 0;
			Integer eurekaTotal = 0;
			Integer ribbonTotal = 0;
			Integer feignTotal = 0;
			Integer zuulTotal = 0;
			Integer kafkaTotal = 0;
			Integer bootTotal = 0;
			Integer sparkTotal = 0;
			Integer dockerTotal = 0;
			Integer jenkinsTotal = 0;
			Integer mavenTotal = 0;
			Integer gitTotal = 0;
			Integer svnTotal = 0;
			Integer html5Total = 0;
			Integer angularTotal = 0;
			Integer scrumTotal = 0;
			Integer manifestoTotal = 0;
			Integer xpTotal = 0;
			Integer safeTotal = 0;
			Integer pmbokTotal = 0;
			Integer cmmiTotal = 0;
			Integer liderancaTotal = 0;
			Integer negociacaoTotal = 0;
			Integer autoTotal = 0;
			Integer empatiaTotal = 0;
			Integer modeloDeNegocioTotal = 0;
			Integer releaseBandeirasTotal = 0;
			Integer cadastroTotal = 0;
			Integer chbTotal = 0;
			Integer convivenciaTotal = 0;
			Integer ficTotal = 0;
			Integer fraudeTotal = 0;
			Integer migracaoTotal = 0;
			Integer monitoracaoTotal = 0;
			Integer precoTotal = 0;
			Integer regulatorioTotal = 0;
			Integer relatorioTotal = 0;
			Integer settlementTotal = 0;
			Integer sieTotal = 0;
			Integer demandasTotal = 0;
			Integer dryTotal = 0;
			Integer kissTotal = 0;
			Integer solidTotal = 0;
			Integer yagniTotal = 0;
			Integer modeloDeDadosTotal = 0;
			Integer umlTotal = 0;
			Integer contadorMedia = 0;

			String role = getCellValue(sh, i, ROLE_CELL_INDEX);

			for (int j = 1; j < sh.getPhysicalNumberOfRows(); j++) {
				if (!BY_ROLE || (getCellValue(sh, j, ROLE_CELL_INDEX).toString() != "#N/D"
						&& role.equals(getCellValue(sh, j, ROLE_CELL_INDEX)))) {
					javaTotal += getNumericCellValue(sh, j, JAVA_CELL_INDEX).intValue();
					jsfTotal += getNumericCellValue(sh, j, JSF_CELL_INDEX).intValue();
					ejbTotal += getNumericCellValue(sh, j, EJB_CELL_INDEX).intValue();
					springTotal += getNumericCellValue(sh, j, SPRING_CELL_INDEX).intValue();
					batchTotal += getNumericCellValue(sh, j, BATCH_CELL_INDEX).intValue();
					jpaTotal += getNumericCellValue(sh, j, JPA_CELL_INDEX).intValue();
					sqlTotal += getNumericCellValue(sh, j, SQL_CELL_INDEX).intValue();
					oracleTotal += getNumericCellValue(sh, j, ORACLE_CELL_INDEX).intValue();
					jmsTotal += getNumericCellValue(sh, j, JMS_CELL_INDEX).intValue();
					microserviceTotal += getNumericCellValue(sh, j, MICROSERVICE_CELL_INDEX).intValue();
					soapTotal += getNumericCellValue(sh, j, SOAP_CELL_INDEX).intValue();
					restTotal += getNumericCellValue(sh, j, REST_CELL_INDEX).intValue();
					eurekaTotal += getNumericCellValue(sh, j, EUREKA_CELL_INDEX).intValue();
					ribbonTotal += getNumericCellValue(sh, j, RIBBON_CELL_INDEX).intValue();
					feignTotal += getNumericCellValue(sh, j, FEIGN_CELL_INDEX).intValue();
					zuulTotal += getNumericCellValue(sh, j, ZUUL_CELL_INDEX).intValue();
					kafkaTotal += getNumericCellValue(sh, j, KAFKA_CELL_INDEX).intValue();
					bootTotal += getNumericCellValue(sh, j, BOOT_CELL_INDEX).intValue();
					sparkTotal += getNumericCellValue(sh, j, SPARK_CELL_INDEX).intValue();
					dockerTotal += getNumericCellValue(sh, j, DOCKER_CELL_INDEX).intValue();
					jenkinsTotal += getNumericCellValue(sh, j, JENKINS_CELL_INDEX).intValue();
					mavenTotal += getNumericCellValue(sh, j, MAVEN_CELL_INDEX).intValue();
					gitTotal += getNumericCellValue(sh, j, GIT_CELL_INDEX).intValue();
					svnTotal += getNumericCellValue(sh, j, SVN_CELL_INDEX).intValue();
					html5Total += getNumericCellValue(sh, j, HTML5_CELL_INDEX).intValue();
					angularTotal += getNumericCellValue(sh, j, ANGULAR_CELL_INDEX).intValue();
					scrumTotal += getNumericCellValue(sh, j, SCRUM_CELL_INDEX).intValue();
					manifestoTotal += getNumericCellValue(sh, j, MANIFESTO_CELL_INDEX).intValue();
					xpTotal += getNumericCellValue(sh, j, XP_CELL_INDEX).intValue();
					safeTotal += getNumericCellValue(sh, j, SAFE_CELL_INDEX).intValue();
					pmbokTotal += getNumericCellValue(sh, j, PMBOK_CELL_INDEX).intValue();
					cmmiTotal += getNumericCellValue(sh, j, CMMI_CELL_INDEX).intValue();
					liderancaTotal += getNumericCellValue(sh, j, LIDERANCA_CELL_INDEX).intValue();
					negociacaoTotal += getNumericCellValue(sh, j, NEGOCIACAO_CELL_INDEX).intValue();
					autoTotal += getNumericCellValue(sh, j, AUTOCONHECIMENTO_CELL_INDEX).intValue();
					empatiaTotal += getNumericCellValue(sh, j, EMPATIA_CELL_INDEX).intValue();
					modeloDeNegocioTotal += getNumericCellValue(sh, j, MODELO_NEGOCIO_CELL_INDEX).intValue();
					releaseBandeirasTotal += getNumericCellValue(sh, j, RELEASEBANDEIRAS_CELL_INDEX).intValue();
					cadastroTotal += getNumericCellValue(sh, j, CADASTRO_CELL_INDEX).intValue();
					chbTotal += getNumericCellValue(sh, j, CHB_CELL_INDEX).intValue();
					convivenciaTotal += getNumericCellValue(sh, j, CONVIVENCIA_CELL_INDEX).intValue();
					ficTotal += getNumericCellValue(sh, j, FIC_CELL_INDEX).intValue();
					fraudeTotal += getNumericCellValue(sh, j, FRAUDE_CELL_INDEX).intValue();
					migracaoTotal += getNumericCellValue(sh, j, MIGRACAO_CELL_INDEX).intValue();
					monitoracaoTotal += getNumericCellValue(sh, j, MONITORACAO_CELL_INDEX).intValue();
					precoTotal += getNumericCellValue(sh, j, PRECO_CELL_INDEX).intValue();
					regulatorioTotal += getNumericCellValue(sh, j, REGULATORIO_CELL_INDEX).intValue();
					relatorioTotal += getNumericCellValue(sh, j, RELATORIO_CELL_INDEX).intValue();
					settlementTotal += getNumericCellValue(sh, j, SETTLEMENT_CELL_INDEX).intValue();
					sieTotal += getNumericCellValue(sh, j, SIE_CELL_INDEX).intValue();
					demandasTotal += getNumericCellValue(sh, j, DEMANDAS_CELL_INDEX).intValue();
					dryTotal += getNumericCellValue(sh, j, DRY_CELL_INDEX).intValue();
					kissTotal += getNumericCellValue(sh, j, KISS_CELL_INDEX).intValue();
					solidTotal += getNumericCellValue(sh, j, SOLID_CELL_INDEX).intValue();
					yagniTotal += getNumericCellValue(sh, j, YAGNI_CELL_INDEX).intValue();
					modeloDeDadosTotal += getNumericCellValue(sh, j, MODELODEDADOS_CELL_INDEX).intValue();
					umlTotal += getNumericCellValue(sh, j, UML_CELL_INDEX).intValue();
					contadorMedia += 1;
				}
			}

			Workbook workbookResult;
			workbookResult = new XSSFWorkbook();
			Sheet shResult = workbookResult.createSheet();
			Row headerRow = shResult.createRow(1);
			Row generalRow = shResult.createRow(2);
			Row individualRow = shResult.createRow(3);

			headerRow.createCell(1).setCellValue("");
			generalRow.createCell(1).setCellValue("Média");
			individualRow.createCell(1).setCellValue(name);

			headerRow.createCell(2).setCellValue("Java (Puro, estrutura de dados, sintaxe ... )?");
			headerRow.createCell(3).setCellValue("JSF?");
			headerRow.createCell(4).setCellValue("EJB?");
			headerRow.createCell(5).setCellValue("Spring Core / Spring Framework?");
			headerRow.createCell(6).setCellValue("Spring Batch?");
			headerRow.createCell(7).setCellValue("JPA?");
			headerRow.createCell(8).setCellValue("SQL (queries básicas)?");
			headerRow.createCell(9).setCellValue("Oracle (funções analíticas e especificidades)?");
			headerRow.createCell(10).setCellValue("JMS (Fila)?");
			headerRow.createCell(11).setCellValue("Microservice?");
			headerRow.createCell(12).setCellValue("SOAP (Exposição de Webservices em JAVA)?");
			headerRow.createCell(13).setCellValue("REST?");
			headerRow.createCell(14).setCellValue("Eureka?");
			headerRow.createCell(15).setCellValue("Ribbon?");
			headerRow.createCell(16).setCellValue("Feign?");
			headerRow.createCell(17).setCellValue("Zuul?");
			headerRow.createCell(18).setCellValue("Kafka?");
			headerRow.createCell(19).setCellValue("Spring Boot?");
			headerRow.createCell(20).setCellValue("Spark?");
			headerRow.createCell(21).setCellValue("Docker?");
			headerRow.createCell(22).setCellValue("Jenkins?");
			headerRow.createCell(23).setCellValue("Maven (Comandos e utilização)?");
			headerRow.createCell(24).setCellValue("GIT?");
			headerRow.createCell(25).setCellValue("SVN?");
			headerRow.createCell(26).setCellValue("HTML5?");
			headerRow.createCell(27).setCellValue("Angular?");
			headerRow.createCell(28).setCellValue("SCRUM?");
			headerRow.createCell(29).setCellValue("Manifesto Ágil?");
			headerRow.createCell(30).setCellValue("XP?");
			headerRow.createCell(31).setCellValue("SAFE?");
			headerRow.createCell(32).setCellValue("PMBOK?");
			headerRow.createCell(33).setCellValue("CMMI?");
			headerRow.createCell(34).setCellValue("Liderança do time?");
			headerRow.createCell(35).setCellValue("Negociação?");
			headerRow.createCell(36).setCellValue("Autoconhecimento?");
			headerRow.createCell(37).setCellValue("Empatia?");
			headerRow.createCell(38).setCellValue("Modelo de Negócio da Cielo");
			headerRow.createCell(39).setCellValue("Release Bandeiras");
			headerRow.createCell(40).setCellValue("Manutenção Cadastral");
			headerRow.createCell(41).setCellValue("Chargeback");
			headerRow.createCell(42).setCellValue("Convivência");
			headerRow.createCell(43).setCellValue("Financeiro e Contábil");
			headerRow.createCell(44).setCellValue("Prevenção a Fraude");
			headerRow.createCell(45).setCellValue("Migração");
			headerRow.createCell(46).setCellValue("Monitoração Funcional");
			headerRow.createCell(47).setCellValue("Preço e Faturamento");
			headerRow.createCell(48).setCellValue("Regulatórios");
			headerRow.createCell(49).setCellValue("Relatório para Clientes");
			headerRow.createCell(50).setCellValue("Settlement (Liquidação)");
			headerRow.createCell(51).setCellValue("SIE");
			headerRow.createCell(52).setCellValue("Tratamento de Demandas");
			headerRow.createCell(53).setCellValue("DRY?");
			headerRow.createCell(54).setCellValue("KISS?");
			headerRow.createCell(55).setCellValue("SOLID?");
			headerRow.createCell(56).setCellValue("YAGNI?");
			headerRow.createCell(57).setCellValue("Modelagem de Dados?");
			headerRow.createCell(58).setCellValue("UML?");

			generalRow.createCell(2).setCellValue(javaTotal / contadorMedia);
			individualRow.createCell(2).setCellValue(getNumericCellValue(sh, i, JAVA_CELL_INDEX).intValue());
			generalRow.createCell(3).setCellValue(jsfTotal / contadorMedia);
			individualRow.createCell(3).setCellValue(getNumericCellValue(sh, i, JSF_CELL_INDEX).intValue());
			generalRow.createCell(4).setCellValue(ejbTotal / contadorMedia);
			individualRow.createCell(4).setCellValue(getNumericCellValue(sh, i, EJB_CELL_INDEX).intValue());
			generalRow.createCell(5).setCellValue(springTotal / contadorMedia);
			individualRow.createCell(5).setCellValue(getNumericCellValue(sh, i, SPRING_CELL_INDEX).intValue());
			generalRow.createCell(6).setCellValue(batchTotal / contadorMedia);
			individualRow.createCell(6).setCellValue(getNumericCellValue(sh, i, BATCH_CELL_INDEX).intValue());
			generalRow.createCell(7).setCellValue(jpaTotal / contadorMedia);
			individualRow.createCell(7).setCellValue(getNumericCellValue(sh, i, JPA_CELL_INDEX).intValue());
			generalRow.createCell(8).setCellValue(sqlTotal / contadorMedia);
			individualRow.createCell(8).setCellValue(getNumericCellValue(sh, i, SQL_CELL_INDEX).intValue());
			generalRow.createCell(9).setCellValue(oracleTotal / contadorMedia);
			individualRow.createCell(9).setCellValue(getNumericCellValue(sh, i, ORACLE_CELL_INDEX).intValue());
			generalRow.createCell(10).setCellValue(jmsTotal / contadorMedia);
			individualRow.createCell(10).setCellValue(getNumericCellValue(sh, i, JMS_CELL_INDEX).intValue());
			generalRow.createCell(11).setCellValue(microserviceTotal / contadorMedia);
			individualRow.createCell(11).setCellValue(getNumericCellValue(sh, i, MICROSERVICE_CELL_INDEX).intValue());
			generalRow.createCell(12).setCellValue(soapTotal / contadorMedia);
			individualRow.createCell(12).setCellValue(getNumericCellValue(sh, i, SOAP_CELL_INDEX).intValue());
			generalRow.createCell(13).setCellValue(restTotal / contadorMedia);
			individualRow.createCell(13).setCellValue(getNumericCellValue(sh, i, REST_CELL_INDEX).intValue());
			generalRow.createCell(14).setCellValue(eurekaTotal / contadorMedia);
			individualRow.createCell(14).setCellValue(getNumericCellValue(sh, i, EUREKA_CELL_INDEX).intValue());
			generalRow.createCell(15).setCellValue(ribbonTotal / contadorMedia);
			individualRow.createCell(15).setCellValue(getNumericCellValue(sh, i, RIBBON_CELL_INDEX).intValue());
			generalRow.createCell(16).setCellValue(feignTotal / contadorMedia);
			individualRow.createCell(16).setCellValue(getNumericCellValue(sh, i, FEIGN_CELL_INDEX).intValue());
			generalRow.createCell(17).setCellValue(zuulTotal / contadorMedia);
			individualRow.createCell(17).setCellValue(getNumericCellValue(sh, i, ZUUL_CELL_INDEX).intValue());
			generalRow.createCell(18).setCellValue(kafkaTotal / contadorMedia);
			individualRow.createCell(18).setCellValue(getNumericCellValue(sh, i, KAFKA_CELL_INDEX).intValue());
			generalRow.createCell(19).setCellValue(bootTotal / contadorMedia);
			individualRow.createCell(19).setCellValue(getNumericCellValue(sh, i, BOOT_CELL_INDEX).intValue());
			generalRow.createCell(20).setCellValue(sparkTotal / contadorMedia);
			individualRow.createCell(20).setCellValue(getNumericCellValue(sh, i, SPARK_CELL_INDEX).intValue());
			generalRow.createCell(21).setCellValue(dockerTotal / contadorMedia);
			individualRow.createCell(21).setCellValue(getNumericCellValue(sh, i, DOCKER_CELL_INDEX).intValue());
			generalRow.createCell(22).setCellValue(jenkinsTotal / contadorMedia);
			individualRow.createCell(22).setCellValue(getNumericCellValue(sh, i, JENKINS_CELL_INDEX).intValue());
			generalRow.createCell(23).setCellValue(mavenTotal / contadorMedia);
			individualRow.createCell(23).setCellValue(getNumericCellValue(sh, i, MAVEN_CELL_INDEX).intValue());
			generalRow.createCell(24).setCellValue(gitTotal / contadorMedia);
			individualRow.createCell(24).setCellValue(getNumericCellValue(sh, i, GIT_CELL_INDEX).intValue());
			generalRow.createCell(25).setCellValue(svnTotal / contadorMedia);
			individualRow.createCell(25).setCellValue(getNumericCellValue(sh, i, SVN_CELL_INDEX).intValue());
			generalRow.createCell(26).setCellValue(html5Total / contadorMedia);
			individualRow.createCell(26).setCellValue(getNumericCellValue(sh, i, HTML5_CELL_INDEX).intValue());
			generalRow.createCell(27).setCellValue(angularTotal / contadorMedia);
			individualRow.createCell(27).setCellValue(getNumericCellValue(sh, i, ANGULAR_CELL_INDEX).intValue());
			generalRow.createCell(28).setCellValue(scrumTotal / contadorMedia);
			individualRow.createCell(28).setCellValue(getNumericCellValue(sh, i, SCRUM_CELL_INDEX).intValue());
			generalRow.createCell(29).setCellValue(manifestoTotal / contadorMedia);
			individualRow.createCell(29).setCellValue(getNumericCellValue(sh, i, MANIFESTO_CELL_INDEX).intValue());
			generalRow.createCell(30).setCellValue(xpTotal / contadorMedia);
			individualRow.createCell(30).setCellValue(getNumericCellValue(sh, i, XP_CELL_INDEX).intValue());
			generalRow.createCell(31).setCellValue(safeTotal / contadorMedia);
			individualRow.createCell(31).setCellValue(getNumericCellValue(sh, i, SAFE_CELL_INDEX).intValue());
			generalRow.createCell(32).setCellValue(pmbokTotal / contadorMedia);
			individualRow.createCell(32).setCellValue(getNumericCellValue(sh, i, PMBOK_CELL_INDEX).intValue());
			generalRow.createCell(33).setCellValue(cmmiTotal / contadorMedia);
			individualRow.createCell(33).setCellValue(getNumericCellValue(sh, i, CMMI_CELL_INDEX).intValue());
			generalRow.createCell(34).setCellValue(liderancaTotal / contadorMedia);
			individualRow.createCell(34).setCellValue(getNumericCellValue(sh, i, LIDERANCA_CELL_INDEX).intValue());
			generalRow.createCell(35).setCellValue(negociacaoTotal / contadorMedia);
			individualRow.createCell(35).setCellValue(getNumericCellValue(sh, i, NEGOCIACAO_CELL_INDEX).intValue());
			generalRow.createCell(36).setCellValue(autoTotal / contadorMedia);
			individualRow.createCell(36)
					.setCellValue(getNumericCellValue(sh, i, AUTOCONHECIMENTO_CELL_INDEX).intValue());
			generalRow.createCell(37).setCellValue(empatiaTotal / contadorMedia);
			individualRow.createCell(37).setCellValue(getNumericCellValue(sh, i, EMPATIA_CELL_INDEX).intValue());
			generalRow.createCell(38).setCellValue(modeloDeNegocioTotal / contadorMedia);
			individualRow.createCell(38).setCellValue(getNumericCellValue(sh, i, MODELO_NEGOCIO_CELL_INDEX).intValue());
			generalRow.createCell(39).setCellValue(releaseBandeirasTotal / contadorMedia);
			individualRow.createCell(39)
					.setCellValue(getNumericCellValue(sh, i, RELEASEBANDEIRAS_CELL_INDEX).intValue());
			generalRow.createCell(40).setCellValue(cadastroTotal / contadorMedia);
			individualRow.createCell(40).setCellValue(getNumericCellValue(sh, i, CADASTRO_CELL_INDEX).intValue());
			generalRow.createCell(41).setCellValue(chbTotal / contadorMedia);
			individualRow.createCell(41).setCellValue(getNumericCellValue(sh, i, CHB_CELL_INDEX).intValue());
			generalRow.createCell(42).setCellValue(convivenciaTotal / contadorMedia);
			individualRow.createCell(42).setCellValue(getNumericCellValue(sh, i, CONVIVENCIA_CELL_INDEX).intValue());
			generalRow.createCell(43).setCellValue(ficTotal / contadorMedia);
			individualRow.createCell(43).setCellValue(getNumericCellValue(sh, i, FIC_CELL_INDEX).intValue());
			generalRow.createCell(44).setCellValue(fraudeTotal / contadorMedia);
			individualRow.createCell(44).setCellValue(getNumericCellValue(sh, i, FRAUDE_CELL_INDEX).intValue());
			generalRow.createCell(45).setCellValue(migracaoTotal / contadorMedia);
			individualRow.createCell(45).setCellValue(getNumericCellValue(sh, i, MIGRACAO_CELL_INDEX).intValue());
			generalRow.createCell(46).setCellValue(monitoracaoTotal / contadorMedia);
			individualRow.createCell(46).setCellValue(getNumericCellValue(sh, i, MONITORACAO_CELL_INDEX).intValue());
			generalRow.createCell(47).setCellValue(precoTotal / contadorMedia);
			individualRow.createCell(47).setCellValue(getNumericCellValue(sh, i, PRECO_CELL_INDEX).intValue());
			generalRow.createCell(48).setCellValue(regulatorioTotal / contadorMedia);
			individualRow.createCell(48).setCellValue(getNumericCellValue(sh, i, REGULATORIO_CELL_INDEX).intValue());
			generalRow.createCell(49).setCellValue(relatorioTotal / contadorMedia);
			individualRow.createCell(49).setCellValue(getNumericCellValue(sh, i, RELATORIO_CELL_INDEX).intValue());
			generalRow.createCell(50).setCellValue(settlementTotal / contadorMedia);
			individualRow.createCell(50).setCellValue(getNumericCellValue(sh, i, SETTLEMENT_CELL_INDEX).intValue());
			generalRow.createCell(51).setCellValue(sieTotal / contadorMedia);
			individualRow.createCell(51).setCellValue(getNumericCellValue(sh, i, SIE_CELL_INDEX).intValue());
			generalRow.createCell(52).setCellValue(demandasTotal / contadorMedia);
			individualRow.createCell(52).setCellValue(getNumericCellValue(sh, i, DEMANDAS_CELL_INDEX).intValue());
			generalRow.createCell(53).setCellValue(dryTotal / contadorMedia);
			individualRow.createCell(53).setCellValue(getNumericCellValue(sh, i, DRY_CELL_INDEX).intValue());
			generalRow.createCell(54).setCellValue(kissTotal / contadorMedia);
			individualRow.createCell(54).setCellValue(getNumericCellValue(sh, i, KISS_CELL_INDEX).intValue());
			generalRow.createCell(55).setCellValue(solidTotal / contadorMedia);
			individualRow.createCell(55).setCellValue(getNumericCellValue(sh, i, SOLID_CELL_INDEX).intValue());
			generalRow.createCell(56).setCellValue(yagniTotal / contadorMedia);
			individualRow.createCell(56).setCellValue(getNumericCellValue(sh, i, YAGNI_CELL_INDEX).intValue());
			generalRow.createCell(57).setCellValue(modeloDeDadosTotal / contadorMedia);
			individualRow.createCell(57).setCellValue(getNumericCellValue(sh, i, MODELODEDADOS_CELL_INDEX).intValue());
			generalRow.createCell(58).setCellValue(umlTotal / contadorMedia);
			individualRow.createCell(58).setCellValue(getNumericCellValue(sh, i, UML_CELL_INDEX).intValue());

			Drawing drawing = shResult.createDrawingPatriarch();
			ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 20, 20);

			Chart chart = drawing.createChart(anchor);

			CTChart ctChart = ((XSSFChart) chart).getCTChart();
			CTPlotArea ctPlotArea = ctChart.getPlotArea();

			CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
			CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
			ctBoolean.setVal(false);
			ctBarChart.addNewBarDir().setVal(STBarDir.COL);

			CTBarSer ctBarSer = ctBarChart.addNewSer();
			CTSerTx ctSerTx = ctBarSer.addNewTx();
			CTStrRef ctStrRef = ctSerTx.addNewStrRef();
			ctStrRef.setF("Sheet0!$B$4");
			ctBarSer.addNewIdx().setVal(1);
			CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
			ctStrRef = cttAxDataSource.addNewStrRef();
			ctStrRef.setF("Sheet0!$C$2:$BG$2");
			CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
			CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
			ctNumRef.setF("Sheet0!$C$4:$BG$4");

			// at least the border lines in Libreoffice Calc ;-)
			XSSFColor colorBar = new XSSFColor(new byte[] { (byte) 153, (byte) 51, (byte) 153 });
			ctBarSer.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(colorBar.getRGB());

			// telling the BarChart that it has axes and giving them Ids
			ctBarChart.addNewAxId().setVal(123456);
			ctBarChart.addNewAxId().setVal(123457);

			CTAreaChart cTAreaChart = ctPlotArea.addNewAreaChart();
			CTBoolean ctBooleanArea = cTAreaChart.addNewVaryColors();
			ctBooleanArea.setVal(false);

			CTAreaSer ctAreaSer = cTAreaChart.addNewSer();

			CTSerTx ctASerTx = ctAreaSer.addNewTx();
			CTStrRef ctAStrRef = ctASerTx.addNewStrRef();
			ctAStrRef.setF("Sheet0!$B$3");
			ctAreaSer.addNewIdx().setVal(1);
			CTAxDataSource cttAAxDataSource = ctAreaSer.addNewCat();
			ctAStrRef = cttAAxDataSource.addNewStrRef();
			ctAStrRef.setF("Sheet0!$C$2:$BG$2");
			CTNumDataSource ctANumDataSource = ctAreaSer.addNewVal();
			CTNumRef ctANumRef = ctANumDataSource.addNewNumRef();
			ctANumRef.setF("Sheet0!$C$3:$BG$3");

			// at least the border lines in Libreoffice Calc ;-)
			XSSFColor color = new XSSFColor(new byte[] { (byte) 204, (byte) 255, (byte) 51 });
			ctAreaSer.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(color.getRGB());

			// telling the BarChart that it has axes and giving them Ids
			cTAreaChart.addNewAxId().setVal(123456);
			cTAreaChart.addNewAxId().setVal(123457);

			// cat axis
			CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
			ctCatAx.addNewAxId().setVal(123456); // id of the cat axis
			CTScaling ctScaling = ctCatAx.addNewScaling();
			ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
			ctCatAx.addNewDelete().setVal(false);
			ctCatAx.addNewAxPos().setVal(STAxPos.B);
			ctCatAx.addNewCrossAx().setVal(123457); // id of the val axis
			ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

			// val axis
			CTValAx ctValAx = ctPlotArea.addNewValAx();
			ctValAx.addNewAxId().setVal(123457); // id of the val axis
			ctScaling = ctValAx.addNewScaling();
			ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
			ctValAx.addNewDelete().setVal(false);
			ctValAx.addNewAxPos().setVal(STAxPos.L);
			ctValAx.addNewCrossAx().setVal(123456); // id of the cat axis
			ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

			// legend
			CTLegend ctLegend = ctChart.addNewLegend();
			ctLegend.addNewLegendPos().setVal(STLegendPos.B);
			ctLegend.addNewOverlay().setVal(false);

			File outputFile = new File(getOutputFileName(outputFolder, name));
			FileOutputStream out = new FileOutputStream(outputFile);
			workbookResult.write(out);
			out.close();
			System.out.println("Finalizando processamento de resultados do(a) " + name);
		}
	}

	private String getOutputFileName(String outputFolder, String name) {
		return (outputFolder == null ? OUTPUT_FOLDER : outputFolder) + name + ".xlsx";
	}

	private String getCellValue(Sheet sh, int i, int cellIndex) {
		try {
			return sh.getRow(i).getCell(cellIndex).getRichStringCellValue().getString();
		} catch (IllegalStateException | NullPointerException e) {
			return "#N/D";
		}
	}

	private Double getNumericCellValue(Sheet sh, int i, int cellIndex) {
		try {
			return sh.getRow(i).getCell(cellIndex).getNumericCellValue();
		} catch (IllegalStateException e) {
			return 0D;
		}
	}

}
