package at.cicero.wordwriter.core;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by scd on 08.07.2016.
 */
public class WWTableContent {
    public Integer col;
    public Integer row;
    private Integer hmerge;
    private Integer vmerge;
    private List<P> content = new ArrayList<P>();
    private List<Tbl> tableInTable = new ArrayList<Tbl>();
    private String topBorder;
    private String bottomBorder;
    private String leftBorder;
    private String rightBorder;

    public String getTopBorder() {
        return topBorder;
    }

    public void setTopBorder(String topBorder) {
        this.topBorder = topBorder;
    }

    public String getBottomBorder() {
        return bottomBorder;
    }

    public void setBottomBorder(String bottomBorder) {
        this.bottomBorder = bottomBorder;
    }

    public String getLeftBorder() {
        return leftBorder;
    }

    public void setLeftBorder(String leftBorder) {
        this.leftBorder = leftBorder;
    }

    public String getRightBorder() {
        return rightBorder;
    }

    public void setRightBorder(String rightBorder) {
        this.rightBorder = rightBorder;
    }


    public WWTableContent(Integer col, Integer row, P content){
        this.col = col;
        this.row = row;
        hmerge = null;
        vmerge = null;
        this.content.add(content);
        tableInTable = null;
        topBorder = WWDefaults.topBorder;
        bottomBorder = WWDefaults.bottomBorder;
        leftBorder = WWDefaults.leftBorder;
        rightBorder = WWDefaults.rightBorder;
    }

    public WWTableContent(Integer col, Integer row, P content, Integer vmerge, Integer hmerge){
        this.col = col;
        this.row = row;
        this.hmerge = hmerge;
        this.vmerge = vmerge;
        this.content.add(content);
        tableInTable = null;
        topBorder = WWDefaults.topBorder;
        bottomBorder = WWDefaults.bottomBorder;
        leftBorder = WWDefaults.leftBorder;
        rightBorder = WWDefaults.rightBorder;
    }

    public WWTableContent(Integer col, Integer row, Tbl tableInTable){
        this.col = col;
        this.row = row;
        hmerge = null;
        vmerge = null;
        this.tableInTable.add(tableInTable);
        topBorder = WWDefaults.topBorder;
        bottomBorder = WWDefaults.bottomBorder;
        leftBorder = WWDefaults.leftBorder;
        rightBorder = WWDefaults.rightBorder;
    }

    public WWTableContent(Integer col, Integer row, Tbl tableInTable, Integer vmerge, Integer hmerge){
        this.col = col;
        this.row = row;
        this.hmerge = hmerge;
        this.vmerge = vmerge;
        this.tableInTable.add(tableInTable);
        topBorder = WWDefaults.topBorder;
        bottomBorder = WWDefaults.bottomBorder;
        leftBorder = WWDefaults.leftBorder;
        rightBorder = WWDefaults.rightBorder;
    }

    public Integer getHmerge(){
        return hmerge;
    }
    public Integer getVmerge(){
        return vmerge;
    }
    public List<P> getContent(){
        return content;
    }
    public List<Tbl> getTableInTable() { return tableInTable; }
}
