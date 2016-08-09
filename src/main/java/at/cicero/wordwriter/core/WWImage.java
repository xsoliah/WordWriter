package at.cicero.wordwriter.core;

import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.*;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.P;
import org.docx4j.wml.R;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;


/**
 * Created by scd on 02.08.2016.
 */
public class WWImage {
    private WordprocessingMLPackage myword;
    private InputStream inputStream;
    private String imageName;
    private String imagePath;
    private Long fileLenght;
    protected final Double maxWidth = 6000000.0;
    protected final Double maxHeight = 6000000.0;
    private Double sizeFactor;
    private Integer offset = 0;
    private final Integer id1 = WWRelationshipIdBuilder.getRelIdCounter();
    private final Integer id2 = WWRelationshipIdBuilder.getRelIdCounter();
    private String fileNameHint;
    private boolean imageIsBehindDocument;
    private boolean isAnchor;
    private boolean isWrapping;
    private Integer vOffset;
    private Integer hOffset;
    private STAlignH hAllignment;
    private STAlignV vAllignment;
    byte[] bytes;


    public WWImage(String imageName, String imagePath, Double sizeFactor, WordprocessingMLPackage myword){
        try {
            this.imageName = imageName;
            this.imagePath = imagePath;
            File file = new File(imagePath + imageName);
            this.myword = myword;
            this.inputStream = new java.io.FileInputStream(file);
            this.fileLenght = file.length();
            this.sizeFactor = sizeFactor;
            this.inputStream = new java.io.FileInputStream(file);
            this.imageIsBehindDocument = false;
            this.isAnchor = false;
            this.isWrapping = false;
            this.vOffset = null;
            this.hOffset = null;
            this.vAllignment = null;
            this.hAllignment = null;
            this.bytes = new byte[fileLenght.intValue()];
        } catch (FileNotFoundException fnfe){
            fnfe.printStackTrace();
        }
    }

    public String getFileNameHint() {
        return fileNameHint;
    }

    public void setFileNameHint(String fileNameHint) {
        this.fileNameHint = fileNameHint;
    }


    public boolean isImageIsBehindDocument() {
        return imageIsBehindDocument;
    }

    public void setImageIsBehindDocument(boolean imageIsBehindDocument) {
        this.imageIsBehindDocument = imageIsBehindDocument;
        if (imageIsBehindDocument)
            isWrapping = false;
    }
    public boolean isAnchor() {
        return isAnchor;
    }

    public void setAnchor(boolean isAnchor) {
        this.isAnchor = isAnchor;
    }
    public boolean isWrapping() {
        return isWrapping;
    }

    public void setWrapping(boolean isWrapping) {
        this.isWrapping = isWrapping;
        if (isWrapping)
            imageIsBehindDocument = false;
    }
    public Integer getvOffset() {
        return vOffset;
    }

    public void setvOffset(Integer vOffset) {
        this.vOffset = vOffset;
    }

    public Integer gethOffset() {
        return hOffset;
    }

    public void sethOffset(Integer hOffset) {
        this.hOffset = hOffset;
    }

    public P buildImage(){
        try {
            org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();

            P p = factory.createP();

            Double height = maxHeight * sizeFactor;
            Double width = maxWidth * sizeFactor;


            Integer numRead = 0;
            while (offset < bytes.length && (numRead = inputStream.read(bytes, offset, bytes.length - offset)) >= 0) {
                offset += numRead;
            }

            inputStream.close();

            return newImage(myword, bytes, fileNameHint, null, id1, id2, height, width);
        } catch (IOException ioe){
            ioe.printStackTrace();
        } catch (Exception e){
            e.printStackTrace();
        }
        return null;
    }

    public P newImage(WordprocessingMLPackage wordMLPackage, byte[] bytes,
                      String filenameHint, String altText, int id1, int id2, Double maxHeight, Double maxWidth) throws Exception {
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
        Inline inline = imagePart.createImageInline( filenameHint, altText, id1, id2, maxHeight.longValue(), maxWidth.longValue(), true);
        org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
        P  p = factory.createP();
        R run = factory.createR();

        p.getParagraphContent().add(run);
        Drawing drawing = factory.createDrawing();
        run.getRunContent().add(drawing);

        if (isAnchor) {
            String anchorXml = XmlUtils.marshaltoString(inline, true, false, Context.jc, Namespaces.NS_WORD12, "anchor", Inline.class);
            org.docx4j.dml.ObjectFactory dmlFactory = new org.docx4j.dml.ObjectFactory();
            org.docx4j.dml.wordprocessingDrawing.ObjectFactory wordDmlFactory = new org.docx4j.dml.wordprocessingDrawing.ObjectFactory();

            Anchor anchor = (Anchor) XmlUtils.unmarshalString(anchorXml, Context.jc, Anchor.class);
            anchor.setSimplePos(dmlFactory.createCTPoint2D());
            anchor.getSimplePos().setX(0L);
            anchor.getSimplePos().setY(0L);
            anchor.setSimplePosAttr(false);
            anchor.setBehindDoc(imageIsBehindDocument);
            anchor.setPositionH(wordDmlFactory.createCTPosH());
            if (hOffset != null) {
                anchor.getPositionH().setPosOffset(hOffset);
            } else if (hAllignment != null){
                anchor.getPositionH().setAlign(STAlignH.RIGHT);
            }
            anchor.getPositionH().setRelativeFrom(STRelFromH.MARGIN);
            anchor.setPositionV(wordDmlFactory.createCTPosV());
            if (vOffset != null) {
                anchor.getPositionV().setPosOffset(vOffset);
            }
            else if (vAllignment != null){
                anchor.getPositionV().setAlign(STAlignV.TOP);
            }
            anchor.getPositionV().setRelativeFrom(STRelFromV.MARGIN);
            if (isWrapping) {
                CTWrapSquare square = wordDmlFactory.createCTWrapSquare();
                square.setWrapText(STWrapText.BOTH_SIDES);
                anchor.setWrapSquare(square);
            } else
                anchor.setWrapNone(wordDmlFactory.createCTWrapNone());
            drawing.getAnchorOrInline().add(anchor);
        }
        else {
            drawing.getAnchorOrInline().add(inline);
        }

        return p;
    }

    public void sethAllignment(STAlignH alignH){
        this.hAllignment = alignH;
    }

    public STAlignH gethAllignment(){
        return this.hAllignment;
    }

    public void setvAllignment(STAlignV alignV){
        this.vAllignment = alignV;
    }

    public STAlignV getvAllignment(){
        return this.vAllignment;
    }

    public String getImageName() {
        return imageName;
    }

    public void setImageName(String imageName) {
        this.imageName = imageName;
    }

    public String getImagePath() {
        return imagePath;
    }

    public void setImagePath(String imagePath) {
        this.imagePath = imagePath;
    }
}
