package at.cicero.wordwriter.core;

import org.docx4j.dml.wordprocessingDrawing.STAlignH;
import org.docx4j.dml.wordprocessingDrawing.STAlignV;
import org.docx4j.wml.HdrFtrRef;
import org.junit.Test;

import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.junit.Test;

import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.*;

/**
 * Created by scd on 01.08.2016.
 */
public class WordWriterTest {
    @Test
    public void createDoc() throws Exception {
        WordWriter test = new WordWriter();

        WWParagraphFormat format = new WWParagraphFormat(12, "U001TLig", "left", 220);
        WWParagraphFormat format1 = new WWParagraphFormat(22, "U001TLig", "left", 220);


        //-- Header for other pages
        /*P secheader = test.createHeaderOrFooterParagraph(Arrays.asList("", "<bu>w</bu>", "<iu>ey</iu>"));
        test.addHeader(secheader, HdrFtrRef.DEFAULT);*/

        //-- Header for first page
        WWImage image = new WWImage("image.jpg", "", 0.2, test.getMyWord());
        image.setAnchor(true);
        image.setImageIsBehindDocument(false);
        //image.sethOffset(2000000);
        //image.setvOffset(1000000);
        image.sethAllignment(STAlignH.LEFT);
        image.setvAllignment(STAlignV.TOP);
        image.setWrapping(true);
        test.getMainPart().addObject(image.buildImage());
        P header = test.createHeaderOrFooterParagraph(Arrays.asList("<biu>hallo</biu>", "hey", "yo<iu>ooooo</iu>o"), format);
        test.addHeader(header, HdrFtrRef.DEFAULT);
        test.addFooter(header, HdrFtrRef.FIRST);

        P aNr = test.createParagraph("Anlage Nr.", format);
        test.addHLine(aNr, "top");
        test.getMainPart().addObject(aNr);

        P zvv = test.createParagraph("zum Vertrag vom", format);
        test.addHLine(zvv, "top");
        //zvv = test.addHLine(zvv, "top");
        test.getMainPart().addObject(zvv);

        P date = test.createParagraph("10.08.2011", format1);
        test.getMainPart().addObject(date);

        P mit = test.createParagraph("mit", format);
        test.addHLine(mit, "top");
        test.getMainPart().addObject(mit);

        P dbp = test.createParagraph("<buui>Deutsche Bank Privat- und Geschäftskunden AG</buui>", format1);
        test.getMainPart().addObject(dbp);

        P lvB = test.createParagraph("<ubi>LV-Bewertungsbestimmungen</ubi>", 32, "U001TLig", "left", null);
        test.getMainPart().addObject(lvB);

        //-- Table
        WWTable t = test.createWWTable(9, 6);
        List<Integer> shadings = Arrays.asList(2, 4, 5);
        t.setColShading(shadings);
        t.setHeaderRows(3);
        List<Integer> integers = Arrays.asList(800, 2800, 1000, 1100, 1200, 1100);
        t.createTableGrid(integers);

        P tableTitle = test.createParagraph("<b>Bewertungstabelle</b>", 24, "Calibri", "left", 3);
        P subTitle = test.createParagraph("(Die Provision wird errechnet aus der gewichteten Beitragssumme und dem Rabattfaktor) (Gültig für die DB PGK AG und die Berliner Bank. Auf Präfixe in den Produktnamen wird verzichtet.)", 20, "U001TLig", "left");
        P tarifNr = test.createParagraph("<b>Tarif-Nr.</b>", 24, "U001TLig", "left");
        P produktName = test.createParagraph("<b>Produktname</b>", 24, "U001TLig", "left");

        t.addCell(0, 0, tableTitle, null, 6);
        t.addCell(0, 0, subTitle);
        t.addCell(1, 0, tarifNr, 1, null);
        t.addCell(1, 1, produktName, 1, null);
        t.addCell(1, 1, test.createParagraph("(Versicherungsart)", format1));
        P textFeld = test.createParagraph("<b>Die Summe der während der Laufzeit zu zahlenden Beiträge (Zahlbeitragssumme), multipliziert mit dem Laufzeitfaktor (LZF), ergibt die gewichtete Beitragssumme.</b>", 14, "U001TLig", "left");
        t.addCell(1, 2, textFeld, null, 4);
        t.addCell(2, 0, "");
        t.addCell(2, 1, "");
        t.addCell(2, 2, test.createParagraph("<b>Rabatt-Faktor</b>", 18, "U001TLig", "center"), null, null);
        t.addCell(2, 3, test.createParagraph("<b>Kostenschieber A</b>", 18, "U001TLig", "center"), null, null);
        t.addCell(2, 4, test.createParagraph("<b>Laufzeit in Jahren</b>", 18, "U001TLig", "center"), null, null);
        t.addCell(2, 5, test.createParagraph("<b>Laufzeitfaktor LZF</b>", 18, "U001TLig", "center"), null, null);

        P ansparRente = test.createParagraph("<b>Anspar-Rente</b>", 24, "U001TLig", "left");
        ansparRente.getContent().add(test.getLineBreak());
        t.addCell(3, 0, "P3001E", null, null);
        t.addCell(3, 1, ansparRente);
        P beitragsNachzahlung = test.createParagraph("gegen laufende Beitragsnachzahlung", 17, "U001TLig", "left");
        beitragsNachzahlung.getContent().add(test.getLineBreak());
        t.addCell(3, 1, beitragsNachzahlung);
        P einzelvertrag = test.createParagraph("Einzelvertrag", 17, "U001TLig", "left");
        einzelvertrag.getContent().add(test.getLineBreak());

        t.addCell(3, 1, einzelvertrag);
        t.addCell(3, 1, test.createParagraph("als Kollektivversicherung Gruppe", 17, "U001TLig", "left"));
        t.addCell(3, 2, "1,000");
        t.addCell(3, 3, "1");
        t.addCell(3, 4, "<b>1-11</b>", 5, null);
        t.addCell(3, 4, "<b>12</b>", 5, null);
        t.addCell(3, 4, "<b>13</b>", 5, null);
        t.addCell(3, 4, "<b>14</b>", 5, null);
        t.addCell(3, 4, "<b>15</b>", 5, null);
        t.addCell(3, 4, "<b>16</b>", 5, null);
        t.addCell(3, 4, "<b>17</b>", 5, null);
        t.addCell(3, 4, "<b>18</b>", 5, null);
        t.addCell(3, 4, "<b>19</b>", 5, null);
        t.addCell(3, 4, "<b>20</b>", 5, null);
        t.addCell(3, 4, "<b>21</b>", 5, null);
        t.addCell(3, 4, "<b>22</b>", 5, null);
        t.addCell(3, 4, "<b>23</b>", 5, null);
        t.addCell(3, 4, "<b>24</b>", 5, null);
        t.addCell(3, 4, "<b>25-35</b>", 5, null);
        t.setTcBorders(3, 4, "single", "single", "double", "single");

        t.addCell(3, 5, "<b>0,86</b>", 5, null);
        t.addCell(3, 5, "<b>0,87</b>", 5, null);
        t.addCell(3, 5, "<b>0,88</b>", 5, null);
        t.addCell(3, 5, "<b>0,89</b>", 5, null);
        t.addCell(3, 5, "<b>0,9</b>", 5, null);
        t.addCell(3, 5, "<b>0,91</b>", 5, null);
        t.addCell(3, 5, "<b>0,92</b>", 5, null);
        t.addCell(3, 5, "<b>0,93</b>", 5, null);
        t.addCell(3, 5, "<b>0,94</b>", 5, null);
        t.addCell(3, 5, "<b>0,95</b>", 5, null);
        t.addCell(3, 5, "<b>0,96</b>", 5, null);
        t.addCell(3, 5, "<b>0,97</b>", 5, null);
        t.addCell(3, 5, "<b>0,98</b>", 5, null);
        t.addCell(3, 5, "<b>0,99</b>", 5, null);
        t.addCell(3, 5, "<b>1,00</b>", 5, null);
        t.setTcBorders(3, 3, "single", "single", "single", "double");

        t.addCell(4, 0, "P300E");
        t.setTcBorders(4, 0, "single", "single", null, null);
        P spezialTarif = test.createParagraph("Spezial-Tarif", 17, "U001TLig", "left");
        spezialTarif.getContent().add(test.getLineBreak());
        t.addCell(4, 1, spezialTarif);
        t.setTcBorders(4, 1, "single", "single", null, "single");
        t.addCell(4, 2, "1,000");
        t.addCell(4, 3, "0,5");
        t.addCell(4, 4, "");
        t.addCell(4, 5, "");

        P zukunftsRente = test.createParagraph("<b>Zukunftsrente</b>", 24, "U001TLig", "left");
        t.addCell(5, 0, "P300E");
        t.setTcBorders(5, 0, "single", "single", null, null);
        t.addCell(5, 1, zukunftsRente);
        t.setTcBorders(5, 1, "single", "single", null, "single");
        P einmalBetrag = test.createParagraph("gegen Einmalbetrag*", 17, "U001TLig", "left");
        einmalBetrag.getContent().add(test.getLineBreak());
        t.addCell(5, 1, einmalBetrag);
        t.addCell(5, 1, einzelvertrag);
        t.addCell(5, 2, "0,900");
        t.addCell(5, 3, "1");
        t.addCell(5, 4, "");
        t.addCell(5, 5, "");


        t.addCell(6, 0, "P3011E");
        t.setTcBorders(6, 0, "single", "single", null, null);
        t.addCell(6, 1, spezialTarif);
        t.setTcBorders(6, 1, "single", "single", null, "single");
        t.addCell(6, 2, "0,900");
        t.addCell(6, 3, "0,5");
        t.addCell(6, 4, "");
        t.addCell(6, 5, "");

        P sofortRente = test.createParagraph("<b>Sofortrente</b>", 24, "U001TLig", "left");
        t.addCell(7, 0, "P3021E");
        t.setTcBorders(7, 0, "single", "single", null, null);
        t.addCell(7, 1, sofortRente);
        t.setTcBorders(7, 1, "single", "single", null, "single");
        t.addCell(7, 1, einmalBetrag);
        t.addCell(7, 1, einzelvertrag);
        t.addCell(7, 2, "0,900");
        t.addCell(7, 3, "1");
        t.addCell(7, 4, "");
        t.addCell(7, 5, "");

        t.addCell(8, 0, "P3021E");
        t.setTcBorders(8, 0, "single", "single", null, null);
        t.addCell(8, 1, spezialTarif);
        t.setTcBorders(8, 1, "single", "single", null, "single");
        t.addCell(8, 2, "0,900");
        t.addCell(8, 3, "0,5");
        t.addCell(8, 4, "");
        t.addCell(8, 5, "");

        t.setTcBorders(3, 3, "nil", "nil", "nil", "double");
        t.setTcBorders(4, 3, "nil", "nil", "nil", "double");
        t.setTcBorders(5, 3, "nil", "nil", "nil", "double");
        t.setTcBorders(6, 3, "nil", "nil", "nil", "double");
        t.setTcBorders(7, 3, "nil", "nil", "nil", "double");
        t.setTcBorders(8, 3, "nil", "single", "nil", "double");

        Tbl x = t.saveTable();
        test.getMainPart().addObject(x);


        //--Table Footer
        P tableFooter = test.createParagraph("Achtung: Bei der Kapital-- und Rentenversicherung gegen laufende Beitragszahlung kommt bei einer abgekürzten Beitragszahlungsdauer zusätzlich ein weiterer Faktor zur Anwendung. Faktorermittlung: beitragsfreie Jahre x 0,022 + 1,0. Bei beitragsfreien Jahren ab 30 Jahren lautet der Faktor immer 1,66.", 18, "U001TLig", "left");
        test.getMainPart().addObject(tableFooter);

        test.addHorizontalLine();
        test.saveDocument();
    }
    @Test
    public void createPlaceholderDocument(){
        WordWriter wordWriter = new WordWriter();
        Map map = new HashMap();
        map.put("##MIETER_X-first_name-surname##", "Dieter Schoass");
        map.put("##MIETER_X-street-house_no##", "Peterwardeinstrasse 8");
        map.put("##MIETER_X-post_code-city##", "9073 Viktring");
        map.put("##amt_annual_old##", "1.234,00");
        map.put("##amt_used##", "123,14");
        map.put("##request_no##", "01923");
        map.put("##MIETOBJEKT-street-house_no##", "Jauerburggasse 17");
        map.put("##MIETOBJEKT-post_code-city##", "8010 Graz");
        map.put("##MIETER_X-salutation-title-surname##", "Manuel Hess");
        map.put("##time_frame_used##", "21.07.2008-12.01.2016");
        map.put("##time_frame_unused##", "12.01.2016-25.07.2016");
        map.put("##amt_annual_reduced##", "1.234,00");
        map.put("##amt_service##", "499,00");
        map.put("##amt_annual_new##", "249,00");
        map.put("##amt_credit_note##", "29,00");
        map.put("##nav_user##", "der Inder");
        map.put("##amt_unused##", "129,00");
        map.put("##rent_deposit_reduced##", "Tarif E46");
        map.put("##rent_deposit##", "Tarif B32");
        map.put("##doc_date##", "25.07.2016");

        try {
            WWPlaceholderSubstitution.renamePlaceholdersFromHashMap("DkkAG.dotx", "substituted.docx", map);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}