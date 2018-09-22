package read;/*
 * @ClassName: DocxTableReader
 * @Description: Read tables in doc or docx files
 * @Author: 0newing
 * @Date: 2018/9/22 14:38
 * @Version: 1.0
 */

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import util.CloseUtils;

import java.io.*;
import java.util.*;


public class DocTableReader
{
    public Map<String, List<String>> readFromDocx(BufferedInputStream docxFileStream)

    {
        boolean readStatus = false;
        Map<String, List<String>> resultMap = new HashMap<>();
        List<String> headList = new ArrayList<>();
        XWPFDocument document = null;
        boolean isTableHead;
        try
        {
            document = new XWPFDocument(docxFileStream);
            Iterator<XWPFTable> it = document.getTablesIterator();
            while (it.hasNext())
            {
                XWPFTable table = it.next();
                isTableHead = true;
                List<XWPFTableRow> rows = table.getRows();
                for (XWPFTableRow row : rows)
                {
                    if (isTableHead)
                    {
                        //第一行表头的处理
                        List<XWPFTableCell> cells = row.getTableCells();
                        for (XWPFTableCell cell : cells)
                        {
                            int n = 1;
                            String head = cell.getText();
                            if (head == null || "".equals(head))
                            {
                                head = "未命名列" + n;
                                n++;
                            }
                            headList.add(head);
                            resultMap.put(head, new ArrayList<>());
                        }
                        isTableHead = false;
                    }
                    else
                    {
                        //处理表内容，确定表头，找到对应数组添加
                        List<XWPFTableCell> cells = row.getTableCells();
                        for (int i = 0; i < headList.size(); i++)
                        {
                            resultMap.get(headList.get(i)).add(cells.get(i).getText());
                        }
                    }
                }
            }
            readStatus = true;
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
        finally
        {
            CloseUtils.closeObject(document);
        }

        if (readStatus)
        {
            return resultMap;
        }
        else
        {
            return null;
        }
    }

    public Map<String, List<String>> readFromDoc(BufferedInputStream docFileStream)
    {
        Map<String, List<String>> resultMap = new HashMap<>();
        //Not finished yet...
        return resultMap;
    }
}