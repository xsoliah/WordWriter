package at.cicero.wordwriter.core;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.Properties;

/**
 * Created by scd on 13.07.2016.
 */
public class WWGetPropertyValues {
    InputStream inputStream;

    public void getPropValues() throws IOException {

        try {
            Properties prop = new Properties();
            String propFileName = "config.properties";

            inputStream = getClass().getClassLoader().getResourceAsStream(propFileName);

            if (inputStream != null) {
                prop.load(inputStream);
            } else {
                throw new FileNotFoundException("property file '" + propFileName + "' not found in the classpath");
            }

            WWDefaults.templateDocumentName = prop.getProperty("TemplateDocumentName");
            WWDefaults.afterSubstitutionDocumentName = prop.getProperty("AfterSubstitutionDocumentName");
            WWDefaults.dotxEnding = prop.getProperty("DotxEnding");
            WWDefaults.docxURI = prop.getProperty("DocxURI");
            WWDefaults.rId = Integer.parseInt(prop.getProperty("RelationshipBuilderStartId"));
            WWDefaults.spacingSetLine = Integer.parseInt(prop.getProperty("SpacingSetLine"));
            WWDefaults.spacingSetLineRule = prop.getProperty("SpacingSetLineRule");
            WWDefaults.pageBreak = prop.getProperty("PageBreak");
            WWDefaults.relationshipIdString = prop.getProperty("RelationshipIdString");
            WWDefaults.wwTableContentSize = Integer.parseInt(prop.getProperty("WWTableContentSize"));
            WWDefaults.wwTableHeaderRows = Integer.parseInt(prop.getProperty("WWTableHeaderRows"));
            WWDefaults.wwTableStyle = prop.getProperty("WWTableStyle");
            WWDefaults.shadingSetColor = prop.getProperty("ShadingSetColor");
            WWDefaults.shadingHeaderFill = prop.getProperty("ShadingHeaderFill");
            WWDefaults.shadingRowFill = prop.getProperty("ShadingRowFill");
            WWDefaults.rsidR = prop.getProperty("RsidR");
            WWDefaults.rsidRDefault = prop.getProperty("RsidRDefault");
            WWDefaults.vMergeRestart = prop.getProperty("VMergeRestart");
            WWDefaults.vMergeContinue = prop.getProperty("VMergeContinue");
            WWDefaults.hMerge = prop.getProperty("HMerge");
            WWDefaults.tableContentOrientation = prop.getProperty("TableContentOrientation");
            WWDefaults.topBorder = prop.getProperty("TopBorder");
            WWDefaults.bottomBorder = prop.getProperty("BottomBorder");
            WWDefaults.leftBorder = prop.getProperty("LeftBorder");
            WWDefaults.rightBorder = prop.getProperty("RightBorder");
            WWDefaults.orientation = prop.getProperty("Orientation");
            WWDefaults.font = prop.getProperty("Font");
            WWDefaults.wwParagraphSize = Integer.parseInt(prop.getProperty("WWParagraphSize"));
            WWDefaults.headerStyle = prop.getProperty("HeaderStyle");
            WWDefaults.orientationRight = prop.getProperty("OrientationRight");
            WWDefaults.orientationCenter = prop.getProperty("OrientationCenter");
            WWDefaults.relativeTo = prop.getProperty("RelativeTo");
            WWDefaults.leader = prop.getProperty("Leader");
            WWDefaults.wpMLLoadError = prop.getProperty("WpMLLoadError");
            WWDefaults.unmarshallError = prop.getProperty("UnMarshallError");
            WWDefaults.uriError = prop.getProperty("URIError");
            WWDefaults.noValueInMapError = prop.getProperty("NoValueInMapError");
            WWDefaults.missingKeyError = prop.getProperty("MissingKeyError");
            WWDefaults.tableContentNull = prop.getProperty("TableContentNull");
            WWDefaults.borderValue = prop.getProperty("BorderValue");
            WWDefaults.borderSize = Integer.parseInt(prop.getProperty("BorderSize"));
            WWDefaults.borderSpace = Integer.parseInt(prop.getProperty("BorderSpace"));
            WWDefaults.borderColor = prop.getProperty("BorderColor");
            WWDefaults.boldRegex = prop.getProperty("BoldRegex");
            WWDefaults.italicRegex = prop.getProperty("ItalicRegex");
            WWDefaults.underlineRegex = prop.getProperty("UnderlineRegex");
            WWDefaults.boldItalicRegex = prop.getProperty("BoldItalicRegex");
            WWDefaults.boldUnderlineRegex = prop.getProperty("BoldUnderlineRegex");
            WWDefaults.italicUnderlineRegex = prop.getProperty("ItalicUnderlineRegex");
            WWDefaults.boldItalicUnderlineRegex = prop.getProperty("BoldItalicUnderlineRegex");

            System.out.print("Default Values saved!\n");
        } catch (Exception e) {
            System.out.println("Exception: " + e);
        } finally {
            inputStream.close();
        }
    }
}
