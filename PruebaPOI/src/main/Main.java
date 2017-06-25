package main;

import java.io.FileNotFoundException;
import java.io.IOException;

import recursos.ModelWordPOI;


public class Main {
	private static final String A1 = "A1", A2 = "A2", A3 = "A3", A4 = "A4", A5 = "A5", A6_1 = "A6.1", A6_2 = "A6.2", 
			A7 = "A7", A8 = "A8", AAA = "AAA", BBB= "BBB";
	/*public static void main(String[] args) {
		try {
			XWPFDocument doc = new XWPFDocument(new FileInputStream(new File("C:\\Users\\Pablo\\Trabajo\\Bankia\\prueba plantilla.docx")));
			List<XWPFTable> table = doc.getTables();
			for (XWPFTable xwpfTable : table) {
				System.out.println(xwpfTable);
				for(int i = 0; i < xwpfTable.getNumberOfRows(); i++){
					for(int j = 0; j < 3; j++){
						if(xwpfTable.getRow(i).getCell(j).getText().equals("Hola")){
							xwpfTable.getRow(i).getCell(j).setText("Cambiado");
							
						}
					}
				}
			}
			FileOutputStream out = new FileOutputStream(new File("C:\\Users\\Pablo\\Trabajo\\Bankia\\out.docx"));
			doc.write(out);
			out.close();
			System.out.println("Correcto");
		    
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}*/
	/*
	public static void main(String[] args){
        String filePath = "C:\\Users\\Pablo\\Trabajo\\Bankia\\prueba plantilla.doc";
        POIFSFileSystem fs = null;        
        try {            
            fs = new POIFSFileSystem(new FileInputStream(filePath));            
            HWPFDocument doc = new HWPFDocument(fs);
            doc = replaceText(doc, "Hola", "Cambiado");
            saveWord(filePath, doc);
        }
        catch(FileNotFoundException e){
            e.printStackTrace();
        }
        catch(IOException e){
            e.printStackTrace();
        }
    }

    private static HWPFDocument replaceText(HWPFDocument doc, String findText, String replaceText){
        Range r1 = doc.getRange(); 

        for (int i = 0; i < r1.numSections(); ++i ) { 
            Section s = r1.getSection(i); 
            for (int x = 0; x < s.numParagraphs(); x++) { 
                Paragraph p = s.getParagraph(x); 
                for (int z = 0; z < p.numCharacterRuns(); z++) { 
                    CharacterRun run = p.getCharacterRun(z); 
                    String text = run.text();
                    if(text.contains(findText)) {
                        run.replaceText(findText, replaceText);
                    } 
                }
            }
        } 
        return doc;
    }

    private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException{
        FileOutputStream out = null;
        try{
            out = new FileOutputStream(new File("C:\\Users\\Pablo\\Trabajo\\Bankia\\out.doc"));
            doc.write(out);
        }
        finally{
            out.close();
        }
    }*/
	
	public static void main(String[] args){
        String filePathIn = "C:\\Users\\Pablo\\Trabajo\\Bankia\\prueba plantilla.docx";
        String filePathOut = "C:\\Users\\Pablo\\Trabajo\\Bankia\\out.docx";
        try {            
            ModelWordPOI modelWordPOI = new ModelWordPOI(filePathIn);
            modelWordPOI.replaceText(AAA, "Texto que sustituye A_AA la primera vez");
            modelWordPOI.replaceText(BBB, "Texto que sustituye B_BB la primera vez");
            modelWordPOI.replaceText(A1, "Texto en A_1 primera vez");
            modelWordPOI.replaceText(A2, "Texto en A_2 primera vez");
            modelWordPOI.replaceText(A3, "Texto en A_3 primera vez");
            modelWordPOI.replaceText(A4, "Texto en A_4 primera vez");
            modelWordPOI.replaceText(A5, "Texto en A_5 primera vez");
            modelWordPOI.replaceText(A6_1, "Texto en A_6.1 primera vez");
            modelWordPOI.replaceText(A6_2, "Texto en A_6.2 primera vez");
            modelWordPOI.replaceText(A7, "Texto en A_7 primera vez");
            modelWordPOI.replaceText(A8, "Texto en A_8 primera vez");
            
            modelWordPOI.aniadeFila(3, 0);
            modelWordPOI.aniadeFila(5, 0);
            
            modelWordPOI.replaceText(A1, "Texto en A_1 segunda vez");
            modelWordPOI.replaceText(A2, "Texto en A_2 segunda vez");
            modelWordPOI.replaceText(A3, "Texto en A_3 segunda vez");
            modelWordPOI.replaceText(A4, "Texto en A_4 segunda vez");
            modelWordPOI.replaceText(A5, "Texto en A_5 segunda vez");
            modelWordPOI.replaceText(A6_1, "Texto en A_6.1 segunda vez");
            modelWordPOI.replaceText(A6_2, "Texto en A_6.2 segunda vez");
            modelWordPOI.replaceText(A7, "Texto en A_7 segunda vez");
            modelWordPOI.replaceText(A8, "Texto en A_8 segunda vez");
            
            modelWordPOI.aniadeFila(1, 0);
            modelWordPOI.aniadeFila(2, 0);
            
            modelWordPOI.replaceText(AAA, "Texto que sustituye A_AA la segunda vez");
            modelWordPOI.replaceText(BBB, "Texto que sustituye B_BB la segunda vez");
            
            modelWordPOI.aniadeFila(3, 0);
            modelWordPOI.aniadeFila(5, 0);
            
            modelWordPOI.replaceText(A1, "Texto en A_1 tercera vez");
            modelWordPOI.replaceText(A2, "Texto en A_2 tercera vez");
            modelWordPOI.replaceText(A3, "Texto en A_3 tercera vez");
            modelWordPOI.replaceText(A4, "Texto en A_4 tercera vez");
            modelWordPOI.replaceText(A5, "Texto en A_5 tercera vez");
            modelWordPOI.replaceText(A6_1, "Texto en A_6.1 tercera vez primera linea");
            modelWordPOI.replaceText(A6_2, "Texto en A_6.2 tercera vez primera linea");
            modelWordPOI.aniadeTextoCelda(1, 5, A6_1);
            modelWordPOI.aniadeTextoCelda(1, 5, A6_2);
            modelWordPOI.replaceText(A6_1, "Texto en A_6.1 tercera vez segunda linea");
            modelWordPOI.replaceText(A6_2, "Texto en A_6.2 tercera vez segunda linea");
            modelWordPOI.replaceText(A7, "Texto en A_7 tercera vez");
            modelWordPOI.replaceText(A8, "Texto en A_8 tercera vez");
            
            modelWordPOI.saveWord(filePathOut);
        }
        catch(FileNotFoundException e){
            e.printStackTrace();
        }
        catch(IOException e){
            e.printStackTrace();
        }
    }

}
