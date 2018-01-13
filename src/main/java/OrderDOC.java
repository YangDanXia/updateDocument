import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.*;
import java.util.Random;

public class OrderDOC {
    public static void main(String[] args){
        File f = new File("词汇表.doc");
        try {
            InputStream in = new FileInputStream(f);
            HWPFDocument ex = new HWPFDocument(in);
            Range range = ex.getRange();
//            将文件内容获取后分行存入数组
            String[] wordList = range.text().split("\r");
            int count = wordList.length;
            String[] newList = new String[count];
            int cbRandCount = 0;// 索引
            int cbPosition;// 位置
            int k =0;
            do {
                Random rand = new Random();
                int r = count - cbRandCount;
                cbPosition = rand.nextInt(r);
                newList[k] = wordList[cbPosition];
                k++;
                cbRandCount++;
                wordList[cbPosition] = wordList[r - 1];// 将最后一位数值赋值给已经被使用的cbPosition
            } while (cbRandCount < count);
            StringBuffer newText = new StringBuffer();
            for(int i=0;i<newList.length;i++){
//                doc文件中换行符为\r
                newText.append(newList[i]).append("\r");
            }
            FileOutputStream os = new FileOutputStream("demo.doc");
            OutputStreamWriter writer = new OutputStreamWriter(os);
            writer.append(newText.toString());
            writer.close();
            os.close();
            in.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
