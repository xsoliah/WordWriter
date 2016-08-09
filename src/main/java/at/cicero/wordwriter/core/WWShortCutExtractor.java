package at.cicero.wordwriter.core;

import org.docx4j.wml.*;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by scd on 28.07.2016.
 */
public class WWShortCutExtractor {
    public static List<R> getParagraphfromString(String s, Integer size, String font){
        ObjectFactory factory = new ObjectFactory();

        // split into substrings
        String[] parts = s.split(("<(/?[b|u|i]*)>"));

        // cast to List<R>
        List<R> list = new ArrayList<R>();
        for (int i = 0; i < parts.length; i++){
            R r = factory.createR();
            RPr rpr = factory.createRPr();
            r.setRPr(rpr);
            HpsMeasure hpsMeasure = factory.createHpsMeasure();
            rpr.setSz(hpsMeasure);
            hpsMeasure.setVal(BigInteger.valueOf(size));
            RFonts f = factory.createRFonts();
            rpr.setRFonts(f);
            f.setAscii(font);
            f.setHAnsi(font);
            Text t = factory.createText();
            r.getContent().add(t);
            t.setValue(parts[i]);
            list.add(r);
        }

        Pattern bPattern = Pattern.compile("<[[u|i]?]*b[[u|i]?]*>(.*?)</[[u|i]?]*b[[u|i]?]*>");
        WWShortCutExtractor.extract(s, bPattern, list, parts);

        Pattern uPattern = Pattern.compile("<[[b|i]?]*u[[b|i]?]*>(.*?)</[[b|i]?]*u[[b|i]?]*>");
        WWShortCutExtractor.extract(s, uPattern, list, parts);

        Pattern iPattern = Pattern.compile("<[[u|b]?]*i[[u|b]?]*>(.*?)</[[u|b]?]*i[[u|b]?]*>");
        WWShortCutExtractor.extract(s, iPattern, list, parts);

        return list;
    }

    public static List<R> getParagraphfromString(String s, WWParagraphFormat format){
        ObjectFactory factory = new ObjectFactory();

        // split into substrings
        String[] parts = s.split(("<(/?[b|u|i]*)>"));

        // cast to List<R>
        List<R> list = new ArrayList<R>();
        for (int i = 0; i < parts.length; i++){
            R r = factory.createR();
            RPr rpr = factory.createRPr();
            r.setRPr(rpr);
            HpsMeasure hpsMeasure = factory.createHpsMeasure();
            rpr.setSz(hpsMeasure);
            hpsMeasure.setVal(BigInteger.valueOf(format.getSize()));
            RFonts f = factory.createRFonts();
            rpr.setRFonts(f);
            f.setAscii(format.getFont());
            f.setHAnsi(format.getFont());
            Text t = factory.createText();
            r.getContent().add(t);
            t.setValue(parts[i]);
            list.add(r);
        }

        Pattern bPattern = Pattern.compile("<[[u|i]?]*b[[u|i]?]*>(.*?)</[[u|i]?]*b[[u|i]?]*>");
        WWShortCutExtractor.extract(s, bPattern, list, parts);

        Pattern uPattern = Pattern.compile("<[[b|i]?]*u[[b|i]?]*>(.*?)</[[b|i]?]*u[[b|i]?]*>");
        WWShortCutExtractor.extract(s, uPattern, list, parts);

        Pattern iPattern = Pattern.compile("<[[u|b]?]*i[[u|b]?]*>(.*?)</[[u|b]?]*i[[u|b]?]*>");
        WWShortCutExtractor.extract(s, iPattern, list, parts);

        return list;
    }

    public static void extract(String s, Pattern pattern, List<R> list, String[] parts){
        ObjectFactory factory = new ObjectFactory();
        Matcher matcher = pattern.matcher(s);
        while (matcher.find()) {
            String text = matcher.group(1);
            for (int i = 0; i < parts.length; i++){
                if (text.equals(parts[i])){
                    RPr rpr = list.get(i).getRPr();
                    if (pattern.toString().equals(WWDefaults.boldRegex)) {
                        rpr.setB(new BooleanDefaultTrue());
                    }
                    else if(pattern.toString().equals(WWDefaults.underlineRegex)){
                        U u = factory.createU();
                        rpr.setU(u);
                        UnderlineEnumeration enumeration = UnderlineEnumeration.SINGLE;
                        u.setVal(enumeration);
                    }
                    else if (pattern.toString().equals(WWDefaults.italicRegex)){
                        rpr.setI(new BooleanDefaultTrue());
                    }
                }
            }
        }
    }
}
