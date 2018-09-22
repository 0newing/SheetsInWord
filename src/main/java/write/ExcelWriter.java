package write;

/*
 * @ClassName: ExcelWriter
 * @Description: Write tables into xls or xlsx files
 * @Author: 0newing
 * @Date: 2018/9/22 14:40
 * @Version: 1.0
 */


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import util.CloseUtils;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.IOException;
import java.util.*;


public class ExcelWriter
{
    public boolean writeToXlsx(Map<String, List<String>> rawContent,
                               BufferedOutputStream outputStream, BufferedInputStream inputStream)
    {
        boolean flag = false;
        Workbook readingFile = null;
        try
        {
            Map<String, List<String>> writingContent = new HashMap<>();
            readingFile = new XSSFWorkbook(inputStream);
            Set<String> headSet = rawContent.keySet();
            Sheet readSheet = readingFile.getSheetAt(0);
            Integer lineCount = null;

            //共有列名，输出列列名
            List<String> headList = new ArrayList<>();

            //确定具体输出列
            Row outRow = readSheet.getRow(0);
            for (Cell cell : outRow)
            {
                cell.setCellType(Cell.CELL_TYPE_STRING);
                String head = cell.getStringCellValue();
                if (headSet.contains(head))
                {
                    headList.add(head);
                    if (lineCount == null)
                    {
                        lineCount = rawContent.get(head).size();
                    }
                }
            }

            //确定具体输出内容
            for (String head : headList)
            {
                writingContent.put(head, rawContent.get(head));
            }

            Workbook writingFile = new XSSFWorkbook();
            Sheet writingSheet = writingFile.createSheet("Sheet1");

            System.out.println(headList);

            //输出
            for (int y = -1; y < lineCount; y++)
            {
                Row row = writingSheet.createRow(y+1);
                for (int x = 0; x < headList.size(); x++)
                {
                    Cell cell = row.createCell(x);
                    if (y == -1)
                    {
                        cell.setCellValue(headList.get(x));
                    }
                    else
                    {
                        cell.setCellValue(writingContent.get(headList.get(x)).get(y));
                    }
                }
            }

            //写文件
            writingFile.write(outputStream);
            flag = true;
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
            CloseUtils.closeObject(readingFile);
        }

        return flag;
    }

    public boolean writeToXls(Map<String, List<String>> content, BufferedOutputStream outputStream,
                              BufferedInputStream inputStream)
    {
        boolean flag = false;
        //Not finished yet...
        return flag;
    }

}