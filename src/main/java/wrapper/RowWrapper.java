package wrapper;

/*
 * @ClassName: RowWrapper
 * @Description: 类描述
 * @Author: 0newing
 * @Date: 2019/4/2 1:39
 * @Version: 1.0
 */


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Iterator;


public class RowWrapper implements Row
{
    private Row rowWrapper;

    public RowWrapper(Row rowWrapper)
    {
        this.rowWrapper = rowWrapper;
    }

    public Row getRowWrapper()
    {
        return rowWrapper;
    }

    @Override
    public Cell createCell(int i)
    {
        return null;
    }

    /**
     * @param i
     * @param i1
     * @deprecated
     */
    @Override
    public Cell createCell(int i, int i1)
    {
        return null;
    }

    @Override
    public Cell createCell(int i, CellType cellType)
    {
        return null;
    }

    @Override
    public void removeCell(Cell cell)
    {

    }

    @Override
    public void setRowNum(int i)
    {

    }

    @Override
    public int getRowNum()
    {
        return 0;
    }

    @Override
    public Cell getCell(int i)
    {
        return null;
    }

    @Override
    public Cell getCell(int i, MissingCellPolicy missingCellPolicy)
    {
        return null;
    }

    @Override
    public short getFirstCellNum()
    {
        return 0;
    }

    @Override
    public short getLastCellNum()
    {
        return 0;
    }

    @Override
    public int getPhysicalNumberOfCells()
    {
        return 0;
    }

    @Override
    public void setHeight(short i)
    {

    }

    @Override
    public void setZeroHeight(boolean b)
    {

    }

    @Override
    public boolean getZeroHeight()
    {
        return false;
    }

    @Override
    public void setHeightInPoints(float v)
    {

    }

    @Override
    public short getHeight()
    {
        return 0;
    }

    @Override
    public float getHeightInPoints()
    {
        return 0;
    }

    @Override
    public boolean isFormatted()
    {
        return false;
    }

    @Override
    public CellStyle getRowStyle()
    {
        return null;
    }

    @Override
    public void setRowStyle(CellStyle cellStyle)
    {

    }

    @Override
    public Iterator<Cell> cellIterator()
    {
        return null;
    }

    @Override
    public Sheet getSheet()
    {
        return null;
    }

    @Override
    public int getOutlineLevel()
    {
        return 0;
    }

    /**
     * Returns an iterator over elements of type {@code T}.
     *
     * @return an Iterator.
     */
    @Override
    public Iterator<Cell> iterator()
    {
        return null;
    }
}