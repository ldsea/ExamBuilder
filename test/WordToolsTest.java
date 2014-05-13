import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;


public class WordToolsTest {

	public static void main(String[] args) throws IOException {
		  XWPFDocument doc = new XWPFDocument();
	      XWPFParagraph p_title = doc.createParagraph(); 
	      p_title.setAlignment(ParagraphAlignment.CENTER);
	      //commit test
	      
	      XWPFParagraph p_single_title = doc.createParagraph();	     
	      XWPFParagraph p_single = doc.createParagraph();	 
	      XWPFParagraph p_muti_title = doc.createParagraph();	
	      XWPFParagraph p_muti = doc.createParagraph();	  
	      XWPFParagraph p_jduge_title = doc.createParagraph();	
	      XWPFParagraph p_jduge = doc.createParagraph();	 	      

	      
	      
	      XWPFRun r1 = p_title.createRun();
	      r1.setText("笔试试卷");
	      r1.setFontFamily("Microsoft Yahei");
	      r1.setFontSize(80);
	      XWPFRun r2 = p1.createRun();
	      r2.setText("测试啊");
	      FileOutputStream out = new FileOutputStream("/Users/ldsea/Desktop/test.docx");
	      doc.write(out);
	      out.close();	      
	}

}
