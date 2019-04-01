package wrapper;

/*
 * @ClassName: CellWrapper
 * @Description: 类描述
 * @Author: 0newing
 * @Date: 2019/4/2 1:39
 * @Version: 1.0
 */


import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Calendar;
import java.util.Date;


public class CellWrapper implements Cell
{
    private Cell cellWrapper;

    public CellWrapper(Cell cellWrapper)
    {
        this.cellWrapper = cellWrapper;
    }

    public Cell getCellWrapper()
    {
        return cellWrapper;
    }

    @Override
    public int getColumnIndex()
    {
        return cellWrapper.getColumnIndex();
    }

    @Override
    public int getRowIndex()
    {
        return cellWrapper.getRowIndex();
    }

    @Override
    public Sheet getSheet()
    {
        return new SheetWrapper(cellWrapper.getSheet());
    }

    @Override
    public Row getRow()
    {
        return new RowWrapper(cellWrapper.getRow());
    }

    /**
     * @param i index
     * @deprecated
     */
    @Override
    public void setCellType(int i)
    {
        cellWrapper.setCellType(i);
    }

    @Override
    public void setCellType(CellType cellType)
    {
        cellWrapper.setCellType(cellType);
    }

    /**
     * @deprecated
     */
    @Override
    public int getCellType()
    {
        return cellWrapper.getCellType();
    }

    @Override
    public CellType getCellTypeEnum()
    {
        return cellWrapper.getCellTypeEnum();
    }

    /**
     * @deprecated
     */
    @Override
    public int getCachedFormulaResultType()
    {
        return cellWrapper.getCachedFormulaResultType();
    }

    @Override
    public CellType getCachedFormulaResultTypeEnum()
    {
        return cellWrapper.getCachedFormulaResultTypeEnum();
    }

    @Override
    public void setCellValue(double v)
    {
        cellWrapper.setCellValue(v);
    }

    @Override
    public void setCellValue(Date date)
    {
        cellWrapper.setCellValue(date);
    }

    @Override
    public void setCellValue(Calendar calendar)
    {
        cellWrapper.setCellValue(calendar);
    }

    @Override
    public void setCellValue(RichTextString richTextString)
    {
        cellWrapper.setCellValue(richTextString);
    }

    @Override
    public void setCellValue(String s)
    {
        cellWrapper.setCellValue(s);
    }

    @Override
    public void setCellFormula(String s)
        throws FormulaParseException
    {
        cellWrapper.setCellFormula(s);
    }

    @Override
    public String getCellFormula()
    {
        return cellWrapper.getCellFormula();
    }

    @Override
    public double getNumericCellValue()
    {
        return cellWrapper.getNumericCellValue();
    }

    @Override
    public Date getDateCellValue()
    {
        return cellWrapper.getDateCellValue();
    }

    @Override
    public RichTextString getRichStringCellValue()
    {
        return cellWrapper.getRichStringCellValue();
    }

    @Override
    public String getStringCellValue()
    {
        return cellWrapper.getStringCellValue();
    }

    @Override
    public void setCellValue(boolean b)
    {
        cellWrapper.setCellValue(b);
    }

    @Override
    public void setCellErrorValue(byte b)
    {
        cellWrapper.setCellErrorValue(b);
    }

    @Override
    public boolean getBooleanCellValue()
    {
        return cellWrapper.getBooleanCellValue();
    }

    @Override
    public byte getErrorCellValue()
    {
        return cellWrapper.getErrorCellValue();
    }

    @Override
    public void setCellStyle(CellStyle cellStyle)
    {
        cellWrapper.setCellStyle(cellStyle);
    }

    @Override
    public CellStyle getCellStyle()
    {
        return cellWrapper.getCellStyle();
    }

    @Override
    public void setAsActiveCell()
    {
        cellWrapper.setAsActiveCell();
    }

    @Override
    public CellAddress getAddress()
    {
        return cellWrapper.getAddress();
    }

    @Override
    public void setCellComment(Comment comment)
    {
        cellWrapper.setCellComment(comment);
    }

    @Override
    public Comment getCellComment()
    {
        return cellWrapper.getCellComment();
    }

    @Override
    public void removeCellComment()
    {
        cellWrapper.removeCellComment();
    }

    @Override
    public Hyperlink getHyperlink()
    {
        return cellWrapper.getHyperlink();
    }

    @Override
    public void setHyperlink(Hyperlink hyperlink)
    {
        cellWrapper.setHyperlink(hyperlink);
    }

    @Override
    public void removeHyperlink()
    {
        cellWrapper.removeHyperlink();
    }

    @Override
    public CellRangeAddress getArrayFormulaRange()
    {
        return cellWrapper.getArrayFormulaRange();
    }

    @Override
    public boolean isPartOfArrayFormulaGroup()
    {
        return cellWrapper.isPartOfArrayFormulaGroup();
    }
}