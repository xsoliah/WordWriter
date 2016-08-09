package at.cicero.wordwriter.core;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.model.listnumbering.*;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.*;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.ObjectFactory;

import javax.xml.bind.JAXBException;
import java.io.*;
import java.math.BigInteger;
import java.util.List;
import java.util.Map;

import static at.cicero.wordwriter.core.WWShortCutExtractor.getParagraphfromString;

/**
 * Created by scd on 04.07.2016.
 */

public class WordWriter {
    ObjectFactory factory = new ObjectFactory();
    private WordprocessingMLPackage myword;
    private MainDocumentPart mainpart;
    private RelationshipsPart relationshipsPart;
    private SectPr sectPr;
    private Logger logger = LogManager.getLogger(WordWriter.class);
    private Integer headerNumber = 1;
    private Integer footerNumber = 1;
    public Hdr header;
    public HeaderPart headerPart;

    //---------------------------------------------------------------------
    // Konstruktor
    //---------------------------------------------------------------------
    public WordWriter() {
        try {
            WWReadConfig.config();
            ObjectFactory factory = new ObjectFactory();
            this.myword = WordprocessingMLPackage.createPackage();
            this.mainpart = myword.getMainDocumentPart();
            this.relationshipsPart = myword.getRelationshipsPart();
            this.sectPr = factory.createSectPr();
            //setAbstractNumbering();
        } catch (InvalidFormatException ife) {
            ife.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    //---------------------------------------------------------------------
    // Getter
    //---------------------------------------------------------------------
    public Integer getHeaderNumber(){
        Integer result = headerNumber;
        headerNumber++;
        return result;
    }

    public Integer getFooterNumber(){
        Integer result = footerNumber;
        footerNumber++;
        return result;
    }

    public MainDocumentPart getMainPart(){
        return this.mainpart;
    }

    public void setMainpart(MainDocumentPart mainPart){
        this.mainpart = mainPart;
    }

    public WordprocessingMLPackage getMyWord(){
        return this.myword;
}

    public RelationshipsPart getRelsPart(){
        return this.relationshipsPart;
    }

    public void setRelsPart(RelationshipsPart relationshipsPart){
        this.relationshipsPart = relationshipsPart;
    }

    public SectPr getSectPr(){
        return this.sectPr;
    }

    public void setSectPr(SectPr sectPr){
        this.sectPr = sectPr;
    }

    //---------------------------------------------------------------------
    // Create Paragraph
    //---------------------------------------------------------------------

    public P createParagraph(String content) throws IOException {
        List<R> list = getParagraphfromString(content, WWDefaults.wwParagraphSize, WWDefaults.font);
        WWParagraph P = new WWParagraph(list, WWDefaults.orientation, null, null);
        return P.getParagraph();
    }
    public P createParagraph(String content, Integer size) throws IOException {
        List<R> list = getParagraphfromString(content, size, WWDefaults.font);
        WWParagraph P = new WWParagraph(list, WWDefaults.orientation, null, null);
        return P.getParagraph();
    }
    public P createParagraph(String content, Integer size, String font) throws IOException {
        List<R> list = getParagraphfromString(content, size, font);
        WWParagraph P = new WWParagraph(list, WWDefaults.orientation, null, null);
        return P.getParagraph();
    }
    public P createParagraph(String content, Integer size, String font, String orientation) throws IOException {
        List<R> list = getParagraphfromString(content, size, font);
        WWParagraph P = new WWParagraph(list, orientation, null, null);
        return P.getParagraph();
    }
    public P createParagraph(String content, Integer size, String font, String orientation, Integer spacing) throws IOException {
        List<R> list = getParagraphfromString(content, size, font);
        WWParagraph P = new WWParagraph(list, orientation, spacing, null);
        return P.getParagraph();
    }
    public P createParagraph(String content, WWParagraphFormat format) throws IOException {
        List<R> list = getParagraphfromString(content, format.getSize(), format.getFont());
        WWParagraph P = new WWParagraph(list, format.getOrientation(), format.getSpacing(), format.getNumId());
        return P.getParagraph();
    }

    //---------------------------------------------------------------------
    // Create Table
    //---------------------------------------------------------------------
    public WWTable createWWTable(Integer rows, Integer cols){
        WWTable table = new WWTable(rows, cols);
        return table;
    }
    //---------------------------------------------------------------------
    // Einfügen Header & Footer
    //---------------------------------------------------------------------


    public void addFooter(P footerContent, HdrFtrRef ref) throws Exception {
        ObjectFactory factory = new ObjectFactory();
        String relId = WWRelationshipIdBuilder.getRelId();

        //-- creates footer.xml
        FooterPart footerPart = new FooterPart();
        footerPart.getContent().add(footerContent);
        footerPart.setPartName(new PartName("/footer" + getFooterNumber().toString()));
        Relationship rel =  mainpart.addTargetPart(footerPart);
        rel.setId(relId);

        //-- creates SecPr in document.xml
        SectPr sp = sectPr;
        if (ref.equals(HdrFtrRef.FIRST)){
            sp.setTitlePg(new BooleanDefaultTrue());
        }
        FooterReference fr =  factory.createFooterReference();
        fr.setId(relId);
        fr.setType(ref);
        sp.getEGHdrFtrReferences().add(fr);
    }

    public void addHeader(P headerContent, HdrFtrRef ref) throws Exception {
        ObjectFactory factory = new ObjectFactory();
        String relId = WWRelationshipIdBuilder.getRelId();

        //-- creates footer.xml
        headerPart = new HeaderPart();
        headerPart.getContent().add(headerContent);
        headerPart.setPartName(new PartName("/header" + getHeaderNumber().toString()));
        Relationship rel =  mainpart.addTargetPart(headerPart);
        rel.setId(relId);

        //-- creates SecPr in document.xml
        SectPr sp = sectPr;
        if (ref.equals(HdrFtrRef.FIRST)) {
            sp.setTitlePg(new BooleanDefaultTrue());
        }
        HeaderReference hr =  factory.createHeaderReference();
        hr.setId(relId);
        hr.setType(ref);
        sp.getEGHdrFtrReferences().add(hr);
    }

    public P createHeaderOrFooterParagraph(List<String> list){
        WWHeaderBuilder headerBuilder = new WWHeaderBuilder(list);
        return headerBuilder.createHeader();
    }

    public P createHeaderOrFooterParagraph(List<String> list, WWParagraphFormat format){
        WWHeaderBuilder headerBuilder = new WWHeaderBuilder(list, format);
        return headerBuilder.createHeader();
    }


    //---------------------------------------------------------------------
    // add - Methoden
    //---------------------------------------------------------------------

    public void addHorizontalLine(){
        ObjectFactory factory = new ObjectFactory();
        P p = mainpart.addParagraphOfText(null);
        PPr pr = factory.createPPr();
        PPrBase.PBdr bd = factory.createPPrBasePBdr();
        CTBorder border = factory.createCTBorder();
        border.setVal(STBorder.fromValue(WWDefaults.borderValue));
        border.setSz(BigInteger.valueOf(WWDefaults.borderSize));
        border.setSpace(BigInteger.valueOf(WWDefaults.borderSpace));
        border.setColor(WWDefaults.borderColor);
        bd.setBottom(border);
        pr.setPBdr(bd);
        p.setPPr(pr);
    }

    public void add(java.util.List<String> list, Integer listId){
        ObjectFactory factory = new ObjectFactory();
        for (int i = 0; i < list.size(); i++){
            P p = factory.createP();
            PPr pr = factory.createPPr();
            R r = factory.createR();
            Text t = factory.createText();
            t.setValue(list.get(i));
            PPrBase.Spacing sp = factory.createPPrBaseSpacing();
            pr.setSpacing(sp);
            sp.setLine(BigInteger.valueOf(WWDefaults.spacingSetLine));
            sp.setLineRule(STLineSpacingRule.fromValue(WWDefaults.spacingSetLineRule));
            PPrBase.NumPr npr = factory.createPPrBaseNumPr();
            pr.setNumPr(npr);
            PPrBase.NumPr.NumId nid = new PPrBase.NumPr.NumId();
            npr.setNumId(nid);
            nid.setVal(BigInteger.valueOf(listId));

            p.getContent().add(pr);
            r.getContent().add(t);
            p.getContent().add(r);

            this.getMainPart().addObject(p);
        }
    }

    //-- Single Line-Break
    public void addLineBreak(){
        P p = new P();
        R r = new R();
        r.getContent().add(new R.Cr());
        p.getContent().add(r);
        this.getMainPart().addObject(p);
    }

    public R getLineBreak(){
        R r = new R();
        r.getContent().add(new R.Cr());
        return r;
    }

    //-- Full Page Break
    public void addPageBreak(){
        P p = new P();
        R r = new R();
        Br br = new Br();
        br.setType(STBrType.fromValue(WWDefaults.pageBreak));
        r.getContent().add(br);
        p.getContent().add(r);
        this.getMainPart().addObject(p);
    }

    //-- add h-line to P
    public P addHLine(P p, String dir){
        ObjectFactory factory = new ObjectFactory();
        List<Object> list = p.getContent();
        PPrBase.PBdr bd = factory.createPPrBasePBdr();
        for (int i = 0; i < list.size(); i++){
            if (list.get(i) instanceof PPr){
                PPr pr = (PPr) list.get(i);
                if (pr.getPBdr() != null)
                    bd = pr.getPBdr();
            }
        }
        CTBorder border = factory.createCTBorder();
        border.setVal(STBorder.fromValue(WWDefaults.borderValue));
        border.setSz(BigInteger.valueOf(WWDefaults.borderSize));
        border.setSpace(BigInteger.valueOf(WWDefaults.borderSpace));
        border.setColor(WWDefaults.borderColor);
        if (dir.equals("top")) {
            bd.setTop(border);
        }
        else if (dir.equals("bottom")) {
            bd.setBottom(border);
        }
        else if (dir.equals("left")) {
            bd.setLeft(border);
        }
        else if (dir.equals("right")) {
            bd.setRight(border);
        }
        for (int i = 0; i < list.size(); i++){
            if (list.get(i) instanceof PPr){
                PPr pr = (PPr) list.get(i);
                pr.setPBdr(bd);
            }
        }
        return p;
    }

    public WWImage createImage(String imageName, String imagePath){
        return new WWImage(imageName, imagePath, 1d, this.getMyWord());
    }

    public WWImage createImage(String imageName, String imagePath, Double sizeFactor){
        return new WWImage(imageName, imagePath, sizeFactor, this.getMyWord());
    }

    public void setAbstractNumbering(){
        /*try {
            NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
            myword.addTargetPart(ndp);
            Numbering.AbstractNum abstractNum = factory.createNumberingAbstractNum();
            abstractNum.setAbstractNumId(BigInteger.valueOf(1));
            Numbering.AbstractNum.MultiLevelType multiLevelType = new Numbering.AbstractNum.MultiLevelType();
            abstractNum.setMultiLevelType(multiLevelType);
            multiLevelType.setVal("singleLevel");
            Lvl lvl = factory.createLvl();
            abstractNum.getLvl().add(lvl);
            lvl.setIlvl(BigInteger.valueOf(0));
            Lvl.Start start = new Lvl.Start();
            lvl.setStart(start);
            start.setVal(BigInteger.valueOf(1));
            NumFmt numFmt = factory.createNumFmt();
            lvl.setNumFmt(numFmt);
            numFmt.setVal(NumberFormat.BULLET);
            Lvl.LvlText text = new Lvl.LvlText();
            lvl.setLvlText(text);
            text.setVal("%1");

            ListNumberingDefinition listNumberingDefinition = new ListNumberingDefinition();

            Numbering.Num num = factory.createNumberingNum();
            num.setNumId(BigInteger.valueOf(3));
            Numbering.Num.AbstractNumId numId = new Numbering.Num.AbstractNumId();
            num.setAbstractNumId(numId);
            numId.setVal(BigInteger.valueOf(1));

            Numbering.Num numbering = factory.createNumberingNum();
            numbering = ndp.addAbstractListNumberingDefinition(abstractNum);
            Numbering.Num.AbstractNumId abstractNumId = new Numbering.Num.AbstractNumId();
            abstractNumId.setVal(BigInteger.valueOf(1));
            numbering.setAbstractNumId(abstractNumId);
            numbering.setNumId(BigInteger.valueOf(3));

        } catch (InvalidFormatException ife){
            ife.printStackTrace();
        }*/
    try {
        NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
        ListHelper listHelper = new ListHelper();
        myword.addTargetPart(ndp);
        Numbering.Num num = listHelper.getUnorderedList(ndp);
        num.setNumId(BigInteger.valueOf(3));
        Numbering.Num.AbstractNumId abstractNumId = new Numbering.Num.AbstractNumId();
        abstractNumId.setVal(BigInteger.valueOf(0));
        num.setAbstractNumId(abstractNumId);
    }catch (InvalidFormatException ife){
        ife.printStackTrace();
    } catch (JAXBException e) {
        e.printStackTrace();
    }
    }

    //---------------------------------------------------------------------
    // Merge-Funktion erstellt docx
    // ---------------------------------------------------------------------
    public void saveDocument() {
        Body body = mainpart.getJaxbElement().getBody();
        body.setSectPr(getSectPr());
        try{
            this.getMyWord().save(new java.io.File(WWDefaults.templateDocumentName));
        } catch (Docx4JException de) {
            de.printStackTrace();
        }
    }
};

/*--Bulleted List
            NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
            myword.getMainDocumentPart().addTargetPart(ndp);
            Numbering.AbstractNum abstractNum = factory.createNumberingAbstractNum();
            abstractNum.setAbstractNumId(BigInteger.valueOf(3));
            Lvl lvl = new Lvl();
            NumFmt numFmt = new NumFmt();
            lvl.setNumFmt(numFmt);
            numFmt.setVal(NumberFormat.fromValue("bullet"));
            abstractNum.getLvl().add(lvl);
            Map m = ndp.getAbstractListDefinitions();
            m.put(abstractNum, m.size());
            Numbering.Num numbering = ndp.addAbstractListNumberingDefinition(abstractNum);
            Numbering.Num.AbstractNumId abstractNumId = new Numbering.Num.AbstractNumId();
            abstractNumId.setVal(BigInteger.valueOf(3));
            numbering.setAbstractNumId(abstractNumId);
            numbering.setNumId(BigInteger.valueOf(3));
*/