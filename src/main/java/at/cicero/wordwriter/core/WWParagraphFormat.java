package at.cicero.wordwriter.core;

/**
 * Created by scd on 03.08.2016.
 */
public class WWParagraphFormat {
    private Integer size;
    private String font;
    private String orientation;
    private Integer numId;
    private Integer spacing;

    public WWParagraphFormat(Integer size, String font, String orientation, Integer spacing, Integer numId){
        this.size = size;
        this.font = font;
        this.orientation = orientation;
        this.spacing = spacing;
        this.numId = numId;
    }

    public WWParagraphFormat(Integer size, String font, String orientation, Integer spacing){
        this.size = size;
        this.font = font;
        this.orientation = orientation;
        this.numId = null;
        this.spacing = spacing;
    }

    public WWParagraphFormat(Integer size, String font, Integer spacing){
        this.size = size;
        this.font = font;
        this.orientation = WWDefaults.orientation;
        this.numId = null;
        this.spacing = spacing;
    }

    public WWParagraphFormat(){
        this.size = WWDefaults.wwParagraphSize;
        this.font = WWDefaults.font;
        this.orientation = WWDefaults.orientation;
        this.numId = null;
        this.spacing = WWDefaults.spacingSetLine;
    }

    public Integer getSize() {
        return size;
    }

    public void setSize(Integer size) {
        this.size = size;
    }

    public String getFont() {
        return font;
    }

    public void setFont(String font) {
        this.font = font;
    }

    public String getOrientation() {
        return orientation;
    }

    public void setOrientation(String orientation) {
        this.orientation = orientation;
    }

    public Integer getNumId() {
        return numId;
    }

    public void setNumId(Integer numId) {
        this.numId = numId;
    }

    public Integer getSpacing() {
        return spacing;
    }

    public void setSpacing(Integer spacing) {
        this.spacing = spacing;
    }
}
