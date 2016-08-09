package at.cicero.wordwriter.core;
import org.docx4j.XmlUtils;
import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.contenttype.CTOverride;
import org.docx4j.openpackaging.contenttype.ContentTypeManager;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import javax.xml.bind.JAXBException;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by scd on 15.07.2016.
 */
public class WWPlaceholderSubstitution {

    public static void renamePlaceholdersFromHashMap(String fileToChange, String fileChanged, Map placeholders) throws IOException {
        try {
            WordprocessingMLPackage wordprocessingMLPackage = WordprocessingMLPackage.load(new java.io.File(fileToChange));

            if (fileToChange.endsWith(WWDefaults.dotxEnding)) {
                ContentTypeManager ctm = wordprocessingMLPackage.getContentTypeManager();
                CTOverride override = ctm.getOverrideContentType().get(new URI(WWDefaults.docxURI));
                override.setContentType(org.docx4j.openpackaging.contenttype.ContentTypes.WORDPROCESSINGML_DOCUMENT);
            }

            MainDocumentPart mainDocumentPart = wordprocessingMLPackage.getMainDocumentPart();

            //Map controller = new HashMap();
            List<String> controller = new ArrayList<String>();
            List<String> missingKeys = new ArrayList<String>();

            Iterator itr = placeholders.entrySet().iterator();
            while (itr.hasNext()){
                Map.Entry pair = (Map.Entry) itr.next();
                controller.add((String) pair.getKey());
            }
            //-- Header- & Footer-Content Substitution
            List<SectionWrapper> sectionWrappers = wordprocessingMLPackage.getDocumentModel().getSections();
            if(sectionWrappers != null) {
                for (int i = 0; i < sectionWrappers.size(); i++) {
                    HeaderFooterPolicy headerFooterPolicy = sectionWrappers.get(i).getHeaderFooterPolicy();
                    HeaderPart headerPart = headerFooterPolicy.getDefaultHeader();
                    FooterPart footerPart = headerFooterPolicy.getDefaultFooter();
                    if (headerPart != null) {
                        for (int j = 0; j < headerPart.getContent().size(); j++) {
                            Object item = headerPart.getContent().get(j);
                            String s = XmlUtils.marshaltoString(item, true);
                            Pattern pattern = Pattern.compile("##(.*?)##");
                            Matcher matcher = pattern.matcher(s);
                            while (matcher.find()) {
                                String toProve = matcher.group(1);
                                String subst = WWPlaceholderSubstitution.eliminateFormats(toProve);
                                if (!placeholders.containsKey("##" + subst + "##")){
                                    missingKeys.add("##" + subst + "##");
                                }
                                s = s.replace(toProve, subst);
                            }
                            Iterator it = placeholders.entrySet().iterator();
                            while (it.hasNext()) {
                                Map.Entry pair = (Map.Entry) it.next();
                                if (s.contains((CharSequence) pair.getKey())) {
                                    controller.remove(pair.getKey());
                                    s = s.replaceAll((String) pair.getKey(), (String) pair.getValue());
                                    headerPart.getContent().set(j, XmlUtils.unmarshalString(s));
                                }
                            }
                        }
                    }
                    if (footerPart != null) {
                        for (int k = 0; k < footerPart.getContent().size(); k++) {
                            Object item = footerPart.getContent().get(k);
                            String s = XmlUtils.marshaltoString(item, true);
                            Pattern pattern = Pattern.compile("##(.*?)##");
                            Matcher matcher = pattern.matcher(s);
                            while (matcher.find()) {
                                String toProve = matcher.group(1);
                                String subst = WWPlaceholderSubstitution.eliminateFormats(toProve);
                                if (!placeholders.containsKey("##" + subst + "##")){
                                    missingKeys.add("##" + subst + "##");
                                }
                                s = s.replace(toProve, subst);
                            }
                            Iterator it = placeholders.entrySet().iterator();
                            while (it.hasNext()) {
                                Map.Entry pair = (Map.Entry) it.next();
                                if (s.contains((CharSequence) pair.getKey())) {
                                    controller.remove(pair.getKey());
                                    s = s.replaceAll((String) pair.getKey(), (String) pair.getValue());
                                    footerPart.getContent().set(k, XmlUtils.unmarshalString(s));
                                }
                            }
                        }
                    }
                }
            }

            //-- Document-Mainpart Substitution
            List<java.lang.Object> list = mainDocumentPart.getContent();
            for (int i = 0; i < list.size(); i++) {
                Object item = list.get(i);
                String s = XmlUtils.marshaltoString(item, true);
                Pattern pattern = Pattern.compile("##(.*?)##");
                Matcher matcher = pattern.matcher(s);
                while (matcher.find()) {
                    String toProve = matcher.group(1);
                    String subst = WWPlaceholderSubstitution.eliminateFormats(toProve);
                    if (!placeholders.containsKey("##" + subst + "##")){
                        //throw new IOException("placeholder not found in map \n");
                        missingKeys.add("##" + subst + "##");
                    }
                    s = s.replace(toProve, subst);
                }
                Iterator it = placeholders.entrySet().iterator();
                while (it.hasNext()){
                    Map.Entry pair = (Map.Entry) it.next();
                    if (s.contains((CharSequence)pair.getKey())){
                        if (pair.getValue() instanceof Tbl && list.get(i) instanceof P){
                            list.set(i, pair.getValue());
                        }
                        else if (pair.getValue() instanceof String){
                            controller.remove(pair.getKey());
                            s = s.replaceAll((String)pair.getKey(), (String)pair.getValue());
                            list.set(i, XmlUtils.unmarshalString(s));
                        }
                    }
                }
            }

            if (!controller.isEmpty()){
                //throw new IOException("Unused Placeholders in Map! \n");
                Iterator it = controller.iterator();
                System.out.print(WWDefaults.noValueInMapError + "\n");
                while (it.hasNext()){
                    System.out.print(it.next().toString() + "\n");
                }
            }

            if (!missingKeys.isEmpty()){
                Iterator it = missingKeys.iterator();
                System.out.print(WWDefaults.missingKeyError + "\n");
                while (it.hasNext()){
                    System.out.print(it.next().toString() + "\n");
                }
            }

            //-- Save to fileChanged
            wordprocessingMLPackage.save(new java.io.File(fileChanged));

        } catch (Docx4JException ex) {
            System.out.print(WWDefaults.wpMLLoadError);
            ex.printStackTrace();
        } catch (JAXBException jex) {
            System.out.print(WWDefaults.unmarshallError);
            jex.printStackTrace();
        } catch (URISyntaxException e) {
            System.out.print(WWDefaults.uriError);
            e.printStackTrace();
        }
    }

    private static String eliminateFormats(String toChange){
        Pattern pattern = Pattern.compile("<(.*?)>");
        Matcher matcher = pattern.matcher(toChange);
        while (matcher.find()) {
            String toEliminate = new String("<" + matcher.group(1) + ">");
            toChange = toChange.replace(toEliminate, "");
        }
        return toChange;
    }
}
