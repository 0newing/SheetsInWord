package test;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class Demo1
{
    public static void main(String[] args)
    {
        String buffer = "";
        String path = "Test.docx";
        InputStream is = null;
        try
        {
            is = new FileInputStream(path);
            XWPFDocument doc = new XWPFDocument(is);
            XWPFWordExtractor ex = new XWPFWordExtractor(doc);
            buffer = ex.getText();
            System.out.println(buffer);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
            if (is != null)
            {
                try
                {
                    is.close();
                }
                catch (IOException e)
                {
                    e.printStackTrace();
                }
                is = null;
            }
        }
    }
}
