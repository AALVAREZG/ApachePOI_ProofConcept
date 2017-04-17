package es.aalvarez.poipoc.poi;



import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;



/**
 * A simple WOrdprocessingML document created by POI XWPF API
 *
 * @author Yegor Kozlov, Adaptado por Antonio Álvarez.
 */

public class GeneraArchivoWord {
	
	
	public String replaceTextFound(String path, int id, String nombre, String texto) throws IOException {
		
		String realContextPath = path; 
		String inputfilepath = realContextPath+"/models"+ "/PLANTILLA.docx";
		String outputfilepath = realContextPath+"/docs/"+id+"_Expediente.docx";
		String relativeOutputfilepath = "/docs/"+id+"_Expediente.docx";
		System.out.println(inputfilepath);
		System.out.println(outputfilepath);
		InputStream fs = new FileInputStream(inputfilepath);
		XWPFDocument doc = new XWPFDocument(fs); 
		
		for (XWPFParagraph p : doc.getParagraphs()) {
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		        	String text = r.getText(0);
		        
		        	if (text != null){
		        		if (text.contains("$ID")) {
		        			text = text.replace("$ID", String.valueOf(id));
		        			r.setText(text, 0);
		        		}if(text.contains("$NOMBRE")) {
		        			text = text.replace("$NOMBRE", nombre);
		        			r.setText(text, 0);
		        		}if(text.contains("$TEXTO")) {
			                text = text.replace("$TEXTO", texto);
			                r.setText(text, 0);
		        		}
		        }
		    }
		}
		
		FileOutputStream out = new FileOutputStream(outputfilepath);
	    doc.write(out);
	    out.close();
	}
		return relativeOutputfilepath;
  }
}



