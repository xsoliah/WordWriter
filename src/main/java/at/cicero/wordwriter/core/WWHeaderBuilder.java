package at.cicero.wordwriter.core;

import org.docx4j.dml.wordprocessingDrawing.ObjectFactory;
import org.docx4j.wml.*;

import java.math.BigInteger;
import java.util.Iterator;
import java.util.List;

/**
 * Created by scd on 04.08.2016.
 */
public class WWHeaderBuilder {
    private List<String> contentList;
    private WWParagraphFormat format;

    public WWHeaderBuilder(List<String> contentList){
        this.contentList = contentList;
        this.format = new WWParagraphFormat(WWDefaults.wwParagraphSize, WWDefaults.font, WWDefaults.orientation, null);
    }

    public WWHeaderBuilder(List<String> contentList, WWParagraphFormat format){
        this.contentList = contentList;
        this.format = format;
    }

    public P createHeader(){
        org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
        P p = factory.createP();

        PPr pr = factory.createPPr();
        p.getContent().add(pr);
        PPrBase.PStyle pStyle = factory.createPPrBasePStyle();
        pr.setPStyle(pStyle);
        pStyle.setVal(WWDefaults.headerStyle);

        Jc jc = factory.createJc();
        pr.setJc(jc);
        jc.setVal(JcEnumeration.fromValue((format == null) ? WWDefaults.orientation : format.getOrientation()));

        RPr rpr = factory.createRPr();
        //- size
        HpsMeasure hpsMeasure = factory.createHpsMeasure();
        rpr.setSz(hpsMeasure);
        hpsMeasure.setVal(BigInteger.valueOf((format == null) ? WWDefaults.wwParagraphSize : format.getSize()));
        //- font
        RFonts fonts = factory.createRFonts();
        rpr.setRFonts(fonts);
        fonts.setHAnsi((format == null) ? WWDefaults.font : format.getFont());
        fonts.setAscii((format == null) ? WWDefaults.font : format.getFont());

        List<R> aseL = WWShortCutExtractor.getParagraphfromString(contentList.get(0), (format == null) ? WWDefaults.wwParagraphSize : format.getSize(), (format == null) ? WWDefaults.font : format.getFont());
        Iterator itL = aseL.iterator();
        while (itL.hasNext()){
            p.getContent().add((R) itL.next());
        }

        if (contentList.size() > 1) {
            R confc = factory.createR();
            p.getContent().add(confc);
            R.Ptab ptabc = factory.createRPtab();
            confc.getContent().add(ptabc);
            STPTabAlignment alignmentc = STPTabAlignment.fromValue(WWDefaults.orientationCenter);
            ptabc.setAlignment(alignmentc);
            STPTabRelativeTo relativeToc = STPTabRelativeTo.fromValue(WWDefaults.relativeTo);
            ptabc.setRelativeTo(relativeToc);
            STPTabLeader leader = STPTabLeader.fromValue(WWDefaults.leader);
            ptabc.setLeader(leader);

            List<R> aseC = WWShortCutExtractor.getParagraphfromString(contentList.get(1), (format == null) ? WWDefaults.wwParagraphSize : format.getSize(), (format == null) ? WWDefaults.font : format.getFont());
            Iterator itC = aseC.iterator();
            while (itC.hasNext()){
                p.getContent().add((R) itC.next());
            }
        }

        if (contentList.size() > 2) {
            R confr = factory.createR();
            p.getContent().add(confr);
            confr.setRPr(rpr);
            R.Ptab ptabr = factory.createRPtab();
            confr.getContent().add(ptabr);
            STPTabAlignment alignmentr = STPTabAlignment.fromValue(WWDefaults.orientationRight);
            ptabr.setAlignment(alignmentr);
            STPTabRelativeTo relativeTor = STPTabRelativeTo.fromValue(WWDefaults.relativeTo);
            ptabr.setRelativeTo(relativeTor);
            STPTabLeader leaderr = STPTabLeader.fromValue(WWDefaults.leader);
            ptabr.setLeader(leaderr);

            List<R> aseR = WWShortCutExtractor.getParagraphfromString(contentList.get(2), (format == null) ? WWDefaults.wwParagraphSize : format.getSize(), (format == null) ? WWDefaults.font : format.getFont());
            Iterator itR = aseR.iterator();
            while (itR.hasNext()){
                p.getContent().add((R) itR.next());
            }
        }
        return p;
    }
}
