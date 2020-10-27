/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package readwritedocx;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 *
 * @author User
 */
public class WriteDoc {
    public static void main(String[] args) throws FileNotFoundException, IOException {
        
        String teks = "Ini Lagi UTS Gais";
        //Penempatan File yang disimpan
        String direktori = "D:\\DATA PENTING\\SEMESTER 5\\PEMOGRAMAN JARINGAN\\UTS\\UTSGAIS.doc";
        XWPFDocument document = new XWPFDocument();
        FileOutputStream out = new FileOutputStream(new File(direktori));
        XWPFParagraph paragraph= document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(teks);
        document.write(out);
        out.close();
        System.out.println("Generate DOC sukses");
    }
}
