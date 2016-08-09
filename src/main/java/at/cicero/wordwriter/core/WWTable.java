package at.cicero.wordwriter.core;
import org.docx4j.jaxb.Context;
import org.docx4j.model.table.TblFactory;
import org.docx4j.wml.*;

import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

import static at.cicero.wordwriter.core.WWShortCutExtractor.getParagraphfromString;

/**
 * Created by scd on 07.07.2016.
 */
public class WWTable {
    private WWTableContent[][] tableContents;
    private Integer rows;
    private Integer cols;
    private TblGrid tg;
    private Integer headerRows;
    private List<Integer> tcWidths;
    private List<Integer> colShading;
    private List<Integer> rowShading;

    public List<Integer> getRowShading() {
        return rowShading;
    }

    public void setRowShading(List<Integer> rowShading) {
        this.rowShading = rowShading;
    }

    public List<Integer> getColShading() {
        return colShading;
    }
    public void setColShading(List<Integer> colShading) {
        this.colShading = colShading;
    }


    public WWTable(Integer rows, Integer cols) {
        this.cols = cols;
        this.rows = rows;
        tableContents = new WWTableContent[rows][cols];
        tg = new TblGrid();
        headerRows = WWDefaults.wwTableHeaderRows;
    }

    public void setHeaderRows(Integer headerRows) {
        this.headerRows = headerRows;
    }

    public TblGrid getTg(){
        return tg;
    }

    public void createTableGrid(List<Integer> widths){
        tcWidths = widths;
        for (Integer i = 0; i < cols; i++) {
            TblGridCol gc = new TblGridCol();
            gc.setW(BigInteger.valueOf(widths.get(i).intValue()));
            tg.getGridCol().add(gc);
        }
    }

    public void setTableContent(Integer row, Integer col, Tbl tableInTable){
        if (tableContents[row][col] == null) {
            tableContents[row][col] = new WWTableContent(row, col, tableInTable);
        }
        else{
            tableContents[row][col].getTableInTable().add(tableInTable);
        }
    }


    public void setTableContent(Integer row, Integer col, Tbl tableInTable, Integer vmerge, Integer hmerge){
        if (tableContents[row][col] == null) {
            tableContents[row][col] = new WWTableContent(row, col, tableInTable, vmerge, hmerge);
        }
        else{
            tableContents[row][col].getTableInTable().add(tableInTable);
        }
    }


    public void setTableContent(Integer row, Integer col, P content){
        if (tableContents[row][col] == null) {
            tableContents[row][col] = new WWTableContent(row, col, content);
        }
        else{
            tableContents[row][col].getContent().add(content);
        }
    }

    public void setTableContent(Integer row, Integer col, P content, Integer vmerge, Integer hmerge){
        if (tableContents[row][col] == null) {
            tableContents[row][col] = new WWTableContent(row, col, content, vmerge, hmerge);
        }
        else{
            tableContents[row][col].getContent().add(content);
        }
    }

    public void addCell(Integer row, Integer col, Tbl tableInTable){
        this.setTableContent(row, col, tableInTable);
    }

    public void addCell(Integer row, Integer col, Tbl tableInTable, Integer vmerge, Integer hmerge){
        this.setTableContent(row, col, tableInTable, vmerge, hmerge);
    }

    public void addCell(Integer row, Integer col, String content) throws IOException {
        List<R> list = getParagraphfromString(content, WWDefaults.wwParagraphSize, WWDefaults.font);
        WWParagraph pb = new WWParagraph(list, WWDefaults.tableContentOrientation, null, null);
        P p = pb.getParagraph();
        this.setTableContent(row, col, p);
    }

    public void addCell(Integer row, Integer col, String content, WWParagraphFormat format) throws IOException {
        List<R> list = getParagraphfromString(content, format.getSize(), format.getFont());
        WWParagraph pb = new WWParagraph(list, format.getOrientation(), format.getSpacing(), format.getNumId());
        P p = pb.getParagraph();
        this.setTableContent(row, col, p);
    }

    public void addCell(Integer row, Integer col, String content, Integer vmerge, Integer hmerge) throws IOException{
        List<R> list = getParagraphfromString(content, WWDefaults.wwParagraphSize, WWDefaults.font);
        WWParagraph pb = new WWParagraph(list, WWDefaults.tableContentOrientation, null, null);
        P p = pb.getParagraph();
        this.setTableContent(row, col, p, vmerge, hmerge);
    }

    public void addCell(Integer row, Integer col, P content){
        this.setTableContent(row, col, content);
    }

    public void addCell(Integer row, Integer col, P content, Integer vmerge, Integer hmerge){
        this.setTableContent(row, col, content, vmerge, hmerge);
    }

    //-- Borders: single, double, nil
    public void setTcBorders(Integer row, Integer col, String top, String bottom, String left, String right){
        if (tableContents[row][col] != null) {
            tableContents[row][col].setTopBorder(top);
            tableContents[row][col].setBottomBorder(bottom);
            tableContents[row][col].setLeftBorder(left);
            tableContents[row][col].setRightBorder(right);
        } else {
            System.out.print(WWDefaults.tableContentNull + "\n");
        }
    }

    public Tbl saveTable(){
        ObjectFactory factory = new ObjectFactory();
        Tbl tabelle = TblFactory.createTable(0, 0, 1);
        tabelle.setTblGrid(this.getTg());
        TblPr tpr = factory.createTblPr();
        tabelle.setTblPr(tpr);
        CTTblPrBase.TblStyle ts = new CTTblPrBase.TblStyle();
        tpr.setTblStyle(ts);
        ts.setVal(WWDefaults.wwTableStyle);
        TblWidth tw = factory.createTblWidth();
        tpr.setTblW(tw);
        tw.setW(BigInteger.valueOf(0));

        CTShd shdModulo = factory.createCTShd();
        shdModulo.setColor(WWDefaults.shadingSetColor);
        shdModulo.setVal(STShd.CLEAR);
        shdModulo.setFill(WWDefaults.shadingRowFill);

        CTShd shdHeader = factory.createCTShd();
        shdHeader.setColor(WWDefaults.shadingSetColor);
        shdHeader.setVal(STShd.CLEAR);
        shdHeader.setFill(WWDefaults.shadingHeaderFill);

        for (Integer j = 0; j < rows; j++) {
            Tr tr = factory.createTr();
            if(j < headerRows){
                TrPr trpr = factory.createTrPr();
                tr.setTrPr(trpr);
                BooleanDefaultTrue bdt = Context.getWmlObjectFactory().createBooleanDefaultTrue();
                trpr.getCnfStyleOrDivIdOrGridBefore().add(Context.getWmlObjectFactory().createCTTrPrBaseTblHeader(bdt));
            }
            for (Integer i = 0; i < cols; i++) {
                Tc tc = factory.createTc();
                TcPr tp = factory.createTcPr();
                CTVerticalJc jc = factory.createCTVerticalJc();
                tp.setVAlign(jc);
                jc.setVal(STVerticalJc.fromValue(WWDefaults.tableContentOrientation));
                TcPrInner.TcBorders borders = factory.createTcPrInnerTcBorders();
                tp.setTcBorders(borders);
                if (tableContents[j][i] != null && tableContents[j][i].getTopBorder() != null){
                    CTBorder border = factory.createCTBorder();
                    border.setVal(STBorder.fromValue(tableContents[j][i].getTopBorder()));
                    setBorderStyle(border);
                    borders.setTop(border);
                }
                if (tableContents[j][i] != null && tableContents[j][i].getBottomBorder() != null){
                    CTBorder border = factory.createCTBorder();
                    border.setVal(STBorder.fromValue(tableContents[j][i].getBottomBorder()));
                    setBorderStyle(border);
                    borders.setBottom(border);
                }
                if (tableContents[j][i] != null && tableContents[j][i].getLeftBorder() != null){
                    CTBorder border = factory.createCTBorder();
                    border.setVal(STBorder.fromValue(tableContents[j][i].getLeftBorder()));
                    setBorderStyle(border);
                    borders.setLeft(border);
                }
                if (tableContents[j][i] != null && tableContents[j][i].getRightBorder() != null){
                    CTBorder border = factory.createCTBorder();
                    border.setVal(STBorder.fromValue(tableContents[j][i].getRightBorder()));
                    setBorderStyle(border);
                    borders.setRight(border);
                }
                if (tcWidths != null) {
                    TblWidth tblWidth = factory.createTblWidth();
                    tp.setTcW(tblWidth);
                    tblWidth.setW(BigInteger.valueOf(tcWidths.get(i)));
                }
                TcPrInner.VMerge vm = factory.createTcPrInnerVMerge();
                //-- hMerge mit GridSpan
                if (tableContents[j][i] != null && tableContents[j][i].getHmerge() != null){
                    TcPrInner.GridSpan gs = factory.createTcPrInnerGridSpan();
                    tp.setGridSpan(gs);
                    gs.setVal(BigInteger.valueOf(tableContents[j][i].getHmerge()));
                }
                //-- vMerge in TcPr
                if (tableContents[j][i] != null && tableContents[j][i].getVmerge() != null) {
                    tp.setVMerge(vm);
                    vm.setVal(WWDefaults.vMergeRestart);
                }
                for(int k = 1; k <= j; k++){
                    if (tableContents[j - k][i] != null && tableContents[j - k][i].getVmerge() != null && tableContents[j - k][i].getVmerge() >= k){
                        tp.setVMerge(vm);
                        vm.setVal(WWDefaults.vMergeContinue);
                    }
                }
                //-- wenn Content == Table
                if (tableContents[j][i] != null && tableContents[j][i].getTableInTable() != null){
                    for (int l = 0; l < tableContents[j][i].getTableInTable().size(); l++) {
                        tc.getContent().add(tableContents[j][i].getTableInTable().get(l));
                    }
                    P p = factory.createP();
                    tc.getContent().add(p);
                    tr.getContent().add(tc);
                    p.setRsidR(WWDefaults.rsidR);
                    p.setRsidRDefault(WWDefaults.rsidRDefault);
                }
                //-- wenn Content == Paragraph
                else if (tableContents[j][i] != null && tableContents[j][i].getContent() != null) {
                    for (int m = 0; m < tableContents[j][i].getContent().size(); m++) {
                        tc.getContent().add(tableContents[j][i].getContent().get(m));
                    }
                    tr.getContent().add(tc);
                }
                //-- Cell-Shading
                if (headerRows != null && j < headerRows) {
                    tp.setShd(shdHeader);
                }
                else  {
                    if (colShading != null && colShading.contains(i)){
                        tp.setShd(shdModulo);
                    }
                    if (rowShading != null && rowShading.contains(j)){
                        tp.setShd(shdModulo);
                    }
                }
                tc.setTcPr(tp);
            }
            tabelle.getContent().add(tr);
        }
        return tabelle;
    }

    protected CTBorder setBorderStyle(CTBorder border){
        border.setColor(WWDefaults.borderColor);
        border.setSz(BigInteger.valueOf(WWDefaults.borderSize));
        border.setSpace(BigInteger.valueOf(WWDefaults.borderSpace));
        return border;
    }
}
