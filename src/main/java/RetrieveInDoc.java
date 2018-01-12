import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class RetrieveInDoc {
    public static void main(String args[]){
        File f = new File("e:/Postgraduate/英语/单词/考研英语大纲5500词汇表 - 副本.doc");
        try {
            InputStream in = new FileInputStream(f);
            HWPFDocument ex = new HWPFDocument(in);
            Range range = ex.getRange();
            Pattern pattern = Pattern.compile("[\\u4e00-\\u9fa5]|\\uff1b|\\uff0c|\\u3001",Pattern.CASE_INSENSITIVE);
            Matcher matcher = pattern.matcher(range.text());
            OutputStream os = new FileOutputStream(f);
            while (matcher.find( )) {
                range.replaceText(matcher.group(),"");
            }
            ex.write(os);
            os.close();
            in.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
