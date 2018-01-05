package word;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hpbf.model.MainContents;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class Docx {
	public static void main(String[] args) {
		readAndWriterTest4();
	}
	public static void readAndWriterTest4()  {
		
        File file = new File("C:/Users/Administrator/Desktop/房源情况.docx");
        String str = "";
        try {
            FileInputStream fis = new FileInputStream(file);
            XWPFDocument xdoc = new XWPFDocument(fis);
            XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
            String doc1 = extractor.getText();
            System.out.println(doc1);
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
