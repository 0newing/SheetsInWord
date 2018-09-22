package test;/*
 * @ClassName Demo2
 * @Description Read tables in docx
 * @Author 0newing
 * @Date 2018/9/22 13:54
 * @Version 1.0
 */

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


public class Demo2
{
    public static void main(String[] args)
    {
        String filePath = "Test.docx";
        BufferedInputStream bufferedInputStream = null;
        try
        {
            bufferedInputStream = new BufferedInputStream(new FileInputStream(new File(filePath)));
            XWPFDocument xwpfDocument = new XWPFDocument(bufferedInputStream);
            Iterator<XWPFTable> it = xwpfDocument.getTablesIterator();
            while (it.hasNext())
            {
                XWPFTable table = it.next();
                List<XWPFTableRow> rows = table.getRows();
                for (XWPFTableRow row : rows)
                {
                    List<XWPFTableCell> cells = row.getTableCells();
                    if (cells instanceof ArrayList){
                        System.out.println("true");
                    }else{
                        System.out.println(cells.getClass());
                    }
                }
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
            if (bufferedInputStream != null)
            {
                try
                {
                    bufferedInputStream.close();
                }
                catch (IOException e)
                {
                    e.printStackTrace();
                }
                bufferedInputStream = null;
            }
        }
    }

}