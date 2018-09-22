package util;

/*
 * @ClassName: CloseUtils
 * @Description: to close those object what implement closeable
 * @Author: 0newing
 * @Date: 2018/9/22 18:51
 * @Version: 1.0
 */


import java.io.Closeable;
import java.io.IOException;


public class CloseUtils
{
    private CloseUtils(){}

    public static void closeObject(Closeable obj){
        if (obj != null){
            try
            {
                obj.close();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
            obj = null;
        }
    }

}