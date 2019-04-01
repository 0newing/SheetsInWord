package wrapper;

/*
 * @ClassName: WorkBookWrapper
 * @Description: 类描述
 * @Author: 0newing
 * @Date: 2019/4/2 1:11
 * @Version: 1.0
 */


import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.Removal;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Spliterator;
import java.util.function.Consumer;


public class WorkBookWrapper implements Closeable, Iterable<Sheet>
{

    private Workbook workbookWrapper;

    public WorkBookWrapper(Workbook workbookWrapper)
    {
        this.workbookWrapper = workbookWrapper;
    }

    public Workbook getWrapper()
    {
        return workbookWrapper;
    }

    int PICTURE_TYPE_EMF = 2;

    int PICTURE_TYPE_WMF = 3;

    int PICTURE_TYPE_PICT = 4;

    int PICTURE_TYPE_JPEG = 5;

    int PICTURE_TYPE_PNG = 6;

    int PICTURE_TYPE_DIB = 7;

    /**
     * @deprecated
     */
    @Deprecated
    @Removal(
        version = "3.18"
    )
    int SHEET_STATE_VISIBLE = 0;

    /**
     * @deprecated
     */
    @Deprecated
    @Removal(
        version = "3.18"
    )
    int SHEET_STATE_HIDDEN = 1;

    /**
     * @deprecated
     */
    @Deprecated
    @Removal(
        version = "3.18"
    )
    int SHEET_STATE_VERY_HIDDEN = 2;

    public int getActiveSheetIndex()
    {
        return workbookWrapper.getActiveSheetIndex();
    }

    public void setActiveSheet(int var1)
    {
        workbookWrapper.setActiveSheet(var1);
    }

    public int getFirstVisibleTab()
    {
        return workbookWrapper.getFirstVisibleTab();
    }

    public void setFirstVisibleTab(int var1)
    {
        workbookWrapper.setFirstVisibleTab(var1);
    }

    public void setSheetOrder(String var1, int var2)
    {
        workbookWrapper.setSheetOrder(var1, var2);
    }

    public void setSelectedTab(int var1)
    {
        workbookWrapper.setSelectedTab(var1);
    }

    public void setSheetName(int var1, String var2)
    {
        workbookWrapper.setSheetName(var1, var2);
    }

    public String getSheetName(int var1)
    {
        return workbookWrapper.getSheetName(var1);
    }

    public int getSheetIndex(String var1)
    {
        return workbookWrapper.getSheetIndex(var1);
    }

    public int getSheetIndex(Sheet var1)
    {
        return workbookWrapper.getSheetIndex(var1);
    }

    public Sheet createSheet()
    {
        return new SheetWrapper(workbookWrapper.createSheet());
    }

    public Sheet createSheet(String var1)
    {
        return new SheetWrapper(workbookWrapper.createSheet(var1));
    }

    public Sheet cloneSheet(int var1)
    {
        return new SheetWrapper(workbookWrapper.cloneSheet(var1));
    }

    public Iterator<Sheet> sheetIterator()
    {
        return workbookWrapper.sheetIterator();
    }

    public int getNumberOfSheets()
    {
        return workbookWrapper.getNumberOfSheets();
    }

    public Sheet getSheetAt(int var1)
    {
        return new SheetWrapper(workbookWrapper.getSheetAt(var1));
    }

    public Sheet getSheet(String var1)
    {
        return new SheetWrapper(workbookWrapper.getSheet(var1));
    }

    public void removeSheetAt(int var1)
    {
        workbookWrapper.removeSheetAt(var1);
    }

    public Font createFont()
    {
        return workbookWrapper.createFont();
    }

    public Font findFont(boolean var1, short var2, short var3, String var4, boolean var5, boolean var6, short var7, byte var8)
    {
        return workbookWrapper.findFont(var1, var2, var3, var4, var5, var6, var7, var8);
    }

    public short getNumberOfFonts()
    {
        return workbookWrapper.getNumberOfFonts();
    }

    public Font getFontAt(short var1)
    {
        return workbookWrapper.getFontAt(var1);
    }

    public CellStyle createCellStyle()
    {
        return workbookWrapper.createCellStyle();
    }

    public int getNumCellStyles()
    {
        return workbookWrapper.getNumCellStyles();
    }

    public CellStyle getCellStyleAt(int var1)
    {
        return workbookWrapper.getCellStyleAt(var1);
    }

    public void write(OutputStream var1)
        throws IOException
    {
        workbookWrapper.write(var1);
    }

    public void close()
        throws IOException
    {
        workbookWrapper.close();
    }

    public int getNumberOfNames()
    {
        return workbookWrapper.getNumberOfNames();
    }

    public Name getName(String var1)
    {
        return workbookWrapper.getName(var1);
    }

    public List<? extends Name> getNames(String var1)
    {
        return workbookWrapper.getNames(var1);
    }

    public List<? extends Name> getAllNames()
    {
        return workbookWrapper.getAllNames();
    }

    public Name getNameAt(int var1)
    {
        return workbookWrapper.getNameAt(var1);
    }

    public Name createName()
    {
        return workbookWrapper.createName();
    }

    public int getNameIndex(String var1)
    {
        return workbookWrapper.getNameIndex(var1);
    }

    public void removeName(int var1)
    {
        workbookWrapper.removeName(var1);
    }

    public void removeName(String var1)
    {
        workbookWrapper.removeName(var1);
    }

    public void removeName(Name var1)
    {
        workbookWrapper.removeName(var1);
    }

    public int linkExternalWorkbook(String var1, Workbook var2)
    {
        return workbookWrapper.linkExternalWorkbook(var1, var2);
    }

    public void setPrintArea(int var1, String var2)
    {
        workbookWrapper.setPrintArea(var1, var2);
    }

    public void setPrintArea(int var1, int var2, int var3, int var4, int var5)
    {
        workbookWrapper.setPrintArea(var1, var2, var3, var4, var5);
    }

    public String getPrintArea(int var1)
    {
        return workbookWrapper.getPrintArea(var1);
    }

    public void removePrintArea(int var1)
    {
        workbookWrapper.removePrintArea(var1);
    }

    public Row.MissingCellPolicy getMissingCellPolicy()
    {
        return workbookWrapper.getMissingCellPolicy();
    }

    public void setMissingCellPolicy(Row.MissingCellPolicy var1)
    {
        workbookWrapper.setMissingCellPolicy(var1);
    }

    public DataFormat createDataFormat()
    {
        return workbookWrapper.createDataFormat();
    }

    public int addPicture(byte[] var1, int var2)
    {
        return workbookWrapper.addPicture(var1, var2);
    }

    public List<? extends PictureData> getAllPictures()
    {
        return workbookWrapper.getAllPictures();
    }

    public CreationHelper getCreationHelper()
    {
        return workbookWrapper.getCreationHelper();
    }

    public boolean isHidden()
    {
        return workbookWrapper.isHidden();
    }

    public void setHidden(boolean var1)
    {
        workbookWrapper.setHidden(var1);
    }

    public boolean isSheetHidden(int var1)
    {
        return workbookWrapper.isSheetHidden(var1);
    }

    public boolean isSheetVeryHidden(int var1)
    {
        return workbookWrapper.isSheetVeryHidden(var1);
    }

    public void setSheetHidden(int var1, boolean var2)
    {
        workbookWrapper.setSheetHidden(var1, var2);
    }

    /**
     * @deprecated
     */
    @Removal(
        version = "3.18"
    )
    public void setSheetHidden(int var1, int var2)
    {
        workbookWrapper.setSheetHidden(var1, var2);
    }

    public SheetVisibility getSheetVisibility(int var1)
    {
        return workbookWrapper.getSheetVisibility(var1);
    }

    public void setSheetVisibility(int var1, SheetVisibility var2)
    {
        workbookWrapper.setSheetVisibility(var1, var2);

    }

    public void addToolPack(UDFFinder var1)
    {
        workbookWrapper.addToolPack(var1);
    }

    public void setForceFormulaRecalculation(boolean var1)
    {
        workbookWrapper.setForceFormulaRecalculation(var1);
    }

    public boolean getForceFormulaRecalculation()
    {
        return workbookWrapper.getForceFormulaRecalculation();
    }

    public SpreadsheetVersion getSpreadsheetVersion()
    {
        return workbookWrapper.getSpreadsheetVersion();
    }

    public int addOlePackage(byte[] var1, String var2, String var3, String var4)
        throws IOException
    {
        return workbookWrapper.addOlePackage(var1, var2, var3, var4);
    }

    /**
     * Returns an iterator over elements of type {@code T}.
     *
     * @return an Iterator.
     */
    @Override
    public Iterator<Sheet> iterator()
    {
        return workbookWrapper.iterator();
    }

    /**
     * Performs the given action for each element of the {@code Iterable}
     * until all elements have been processed or the action throws an
     * exception.  Unless otherwise specified by the implementing class,
     * actions are performed in the order of iteration (if an iteration order
     * is specified).  Exceptions thrown by the action are relayed to the
     * caller.
     *
     * @param action The action to be performed for each element
     * @throws NullPointerException if the specified action is null
     * @implSpec <p>The default implementation behaves as if:
     * <pre>{@code
     *     for (T t : this)
     *         action.accept(t);
     * }</pre>
     * @since 1.8
     */
    @Override
    public void forEach(Consumer<? super Sheet> action)
    {
        workbookWrapper.forEach(action);
    }

    /**
     * Creates a {@link Spliterator} over the elements described by this
     * {@code Iterable}.
     *
     * @return a {@code Spliterator} over the elements described by this
     * {@code Iterable}.
     * @implSpec The default implementation creates an
     * <em><a href="Spliterator.html#binding">early-binding</a></em>
     * spliterator from the iterable's {@code Iterator}.  The spliterator
     * inherits the <em>fail-fast</em> properties of the iterable's iterator.
     * @implNote The default implementation should usually be overridden.  The
     * spliterator returned by the default implementation has poor splitting
     * capabilities, is unsized, and does not report any spliterator
     * characteristics. Implementing classes can nearly always provide a
     * better implementation.
     * @since 1.8
     */
    @Override
    public Spliterator<Sheet> spliterator()
    {
        return workbookWrapper.spliterator();
    }
}