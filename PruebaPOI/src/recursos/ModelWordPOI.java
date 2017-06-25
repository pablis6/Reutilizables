package recursos;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

public class ModelWordPOI {
	XWPFDocument document;
	String filePathPlantilla;
	public ModelWordPOI(String filePath) {
		try {
			filePathPlantilla = filePath;
			document = new XWPFDocument(new FileInputStream(filePath));
		}  catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public void saveWord(String filePath) throws FileNotFoundException, IOException{
        FileOutputStream out = null;
        try{
            out = new FileOutputStream(new File(filePath));
            document.write(out);
            Desktop.getDesktop().open(new File(filePath));
        }
        finally{
            out.close();
        }
    }
	
	public void replaceText(String findText, String newText){
		//cabecera del documento
		for(XWPFHeader header : document.getHeaderList()){	
			for (XWPFParagraph parrafo : header.getParagraphs()) {
			    List<XWPFRun> runs = parrafo.getRuns();
			    if (runs != null) {
			        for (XWPFRun r : runs) {
			            String text = r.getText(0);
			            if (text != null && text.contains(findText)) {
			                text = text.replace(findText, newText);
			                r.setText(text, 0);
			            }
			        }
			    }
			}
		}
		//cuerpo del documento
		for (XWPFParagraph parrafo : document.getParagraphs()) {
		    List<XWPFRun> runs = parrafo.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            if (text != null && text.contains(findText)) {
		                text = text.replace(findText, newText);
		                r.setText(text, 0);
		            }
		        }
		    }
		}
		//tabla del documento
		for (XWPFTable tbl : document.getTables()) {
			for (XWPFTableRow row : tbl.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					for (XWPFParagraph parrafo : cell.getParagraphs()) {
						for (XWPFRun r : parrafo.getRuns()) {
							String text = r.getText(0);
							if (text != null && text.contains(findText)) {
								text = text.replace(findText, newText);
								r.setText(text, 0);
							}
						}
					}
				}
			}
		}
		//pie del documento
		for(XWPFFooter footer : document.getFooterList()){	
			for (XWPFParagraph parrafo : footer.getParagraphs()) {
			    List<XWPFRun> runs = parrafo.getRuns();
			    if (runs != null) {
			        for (XWPFRun r : runs) {
			            String text = r.getText(0);
			            if (text != null && text.contains(findText)) {
			                text = text.replace(findText, newText);
			                r.setText(text, 0);
			            }
			        }
			    }
			}
		}
	}
	
	public void aniadeFila(int filaPlantilla, int filaDestino){
		try {
			XWPFTable tbl = document.getTables().get(0);
			//abrimos de nuevo la plantilla
			XWPFDocument docAux = new XWPFDocument(new FileInputStream(filePathPlantilla));
			//copiamos la fila de la plantilla que se quiere clonar
			XWPFTableRow filaClonar = docAux.getTables().get(0).getRow(filaPlantilla);
			//creamos una nueva fila (1 x 1)
			XWPFTableRow filaNueva = tbl.insertNewTableRow(tbl.getNumberOfRows() - filaDestino);
			//añadimos las columnas necesarias hasta llegar a ser igual la nueva y la que queremos clonar
			for(int i = 0; i < filaClonar.getTableCells().size(); i++){
				XWPFTableCell celdaNueva = filaNueva.addNewTableCell();
				XWPFTableCell celdaClonar = filaClonar.getCell(i);
				CTTcPr filaNuevaProperties = celdaNueva.getCTTc().addNewTcPr();
				CTTcPr filaClonarProperties = celdaClonar.getCTTc().getTcPr();
				//gridspan
				filaNuevaProperties.addNewGridSpan();
				filaNuevaProperties.setGridSpan(filaClonarProperties.getGridSpan());
				//bordes de celda
				CTTcBorders bordes = filaNuevaProperties.addNewTcBorders();
				//arriba
				bordes.addNewTop();
				filaNuevaProperties.getTcBorders().getTop().setVal(filaClonarProperties.getTcBorders().getTop().getVal());
				//derecha
				bordes.addNewRight();
				filaNuevaProperties.getTcBorders().getRight().setVal(filaClonarProperties.getTcBorders().getRight().getVal());
				//abajo
				bordes.addNewBottom();
				filaNuevaProperties.getTcBorders().getBottom().setVal(filaClonarProperties.getTcBorders().getBottom().getVal());
				///izquierda
				bordes.addNewLeft();
				filaNuevaProperties.getTcBorders().getLeft().setVal(filaClonarProperties.getTcBorders().getLeft().getVal());
				//color
				celdaNueva.setColor(celdaClonar.getColor());
				for(int j = 0; j < celdaClonar.getParagraphs().size(); j++){
					if(j > 0){
						celdaNueva.addParagraph();
					}
					XWPFParagraph parrafoObjetivo = celdaNueva.getParagraphs().get(j);
					XWPFParagraph parrafoClonar = celdaClonar.getParagraphs().get(j);
					for(int k = 0; k < parrafoClonar.getRuns().size(); k++){
						XWPFRun run = parrafoObjetivo.createRun();
						run.setText(parrafoClonar.getText());
					}
					if(parrafoClonar.getRuns().size() == 0 && filaNueva.getTableCells().size() == 1){
						filaNueva.getCtRow().addNewTrPr();
						filaNueva.getCtRow().getTrPr().addNewTrHeight();
						filaNueva.getCtRow().getTrPr().setTrHeightArray(filaClonar.getCtRow().getTrPr().getTrHeightArray());
					}
					
				}
				
			}
			
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public void aniadeTextoCelda(int filaCambiar, int columnaCambiar, String textoAniadir) {
		XWPFTable tbl = document.getTables().get(0);
		XWPFTableRow fila = tbl.getRow(tbl.getNumberOfRows()-filaCambiar);
		XWPFTableCell celda = fila.getCell(columnaCambiar);
		celda.addParagraph().createRun().setText(textoAniadir);
	}
}
