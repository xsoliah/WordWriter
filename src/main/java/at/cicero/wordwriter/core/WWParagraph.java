package at.cicero.wordwriter.core;
import org.docx4j.wml.*;

import java.io.IOException;
import java.math.BigInteger;
import java.util.Iterator;
import java.util.List;

/**
 * Created by scd on 12.07.2016.
 */
public class WWParagraph {
    private List<R> contentList;
    private String orientation;
    private Integer numId;
    private Integer spacing;

    public WWParagraph(List<R> content, String orientation, Integer spacing, Integer numId) {
        this.contentList = content;
        this.orientation = orientation;
        this.numId = numId;
        this.spacing = spacing;
    }

    public void setNumId(Integer numId){
        this.numId = numId;
    }

    public P getParagraph(){
        ObjectFactory factory = new ObjectFactory();
        P p = factory.createP();
        PPr pr = factory.createPPr();
        p.getContent().add(pr);
        PPrBase.Spacing sp = factory.createPPrBaseSpacing();
        pr.setSpacing(sp);
        if (spacing != null)
            sp.setLine(BigInteger.valueOf(spacing));
        else
            sp.setLine(BigInteger.valueOf(WWDefaults.spacingSetLine));
        sp.setLineRule(STLineSpacingRule.fromValue(WWDefaults.spacingSetLineRule));
        if (orientation != null){
            Jc j = factory.createJc();
            pr.setJc(j);
            j.setVal(JcEnumeration.fromValue(orientation));
        }
        if (numId != null){
            PPrBase.NumPr numPr = factory.createPPrBaseNumPr();
            pr.setNumPr(numPr);
            PPrBase.NumPr.NumId nid = factory.createPPrBaseNumPrNumId();
            numPr.setNumId(nid);
            nid.setVal(BigInteger.valueOf(numId));
        }
        if (contentList != null) {
            Iterator it = contentList.iterator();
            while (it.hasNext()){
                p.getContent().add(it.next());
            }
        }
        return p;
    }
}
