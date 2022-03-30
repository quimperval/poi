package org.apache.poi.examples.xwpf.usermodel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

public class SimpleMultiSectionDocument {

    public static void main(String[] args) throws IOException {
        try (XWPFDocument doc = new XWPFDocument()) {
        	
        	
        	for(int i=0; i<3; i++) {
        		XWPFParagraph p = doc.createParagraph();
                XWPFRun run = p.createRun();
                run.setText("Content of section " + i);
                CTP ctp = p.getCTP();
                if(!ctp.isSetPPr()){
                	ctp.addNewPPr();
                }
                
                if(!ctp.getPPr().isSetSectPr()) {
                	ctp.getPPr().addNewSectPr();
                }
                
                if(!ctp.getPPr().getSectPr().isSetPgSz()) {
                	ctp.getPPr().getSectPr().addNewPgSz();
                }
                
                CTPageSz pgSize = ctp.getPPr().getSectPr().getPgSz();
                //Setting orientation as portrait to odds and landscape to even pages
                if(i%2==0) {
                	//A4 = 595x842 / multiply 20 since BigInteger represents 1/20 Point
                	pgSize.setOrient(STPageOrientation.LANDSCAPE);
                	pgSize.setW(BigInteger.valueOf(16840));
                	pgSize.setH(BigInteger.valueOf(11900));
                } else {
                	pgSize.setOrient(STPageOrientation.PORTRAIT);
                	pgSize.setW(BigInteger.valueOf(11900));
                	pgSize.setH(BigInteger.valueOf(16840));
                }
                
        	}
            
            
            try (OutputStream os = new FileOutputStream(new File("simple_multi_section_document.docx"))) {
                doc.write(os);
            }
        }
    }
}
