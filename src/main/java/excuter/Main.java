package excuter;

/*
 * @ClassName: Main
 * @Description: 类描述
 * @Author: 0newing
 * @Date: 2018/9/22 18:17
 * @Version: 1.0
 */


import read.DocTableReader;
import util.CloseUtils;
import write.ExcelWriter;

import java.io.*;
import java.util.List;
import java.util.Map;


public class Main
{
    public static void main(String[] args)
    {
        BufferedInputStream doc = null;
        BufferedInputStream module = null;
        BufferedOutputStream outXlsx = null;
        try
        {
            doc = new BufferedInputStream(new FileInputStream(new File("Test.docx")));
            module = new BufferedInputStream(new FileInputStream(new File("Test.xlsx")));
            outXlsx = new BufferedOutputStream(new FileOutputStream(new File("Result.xlsx")));
            DocTableReader docTableReader = new DocTableReader();
            Map<String, List<String>> map = docTableReader.readFromDocx(doc);
            ExcelWriter excelWriter = new ExcelWriter();
            boolean suc = excelWriter.writeToXlsx(map, outXlsx, module);
            if (suc)
            {
                System.out.println("成功！");
            }
            else
            {
                System.out.println("失败！");
            }
        }
        catch (FileNotFoundException e)
        {
            e.printStackTrace();
        }
        finally
        {
            CloseUtils.closeObject(outXlsx);
            CloseUtils.closeObject(module);
            CloseUtils.closeObject(doc);
        }
    }
}