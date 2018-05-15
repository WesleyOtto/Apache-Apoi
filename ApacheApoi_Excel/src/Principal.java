/*

#License

Copyright (c) 2018 Wesley Otto Garcia Utsuomiya 

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

*/

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;

public class Principal {

	public static void main(String[] args) throws IOException {

		// Criando o arquivo de escrita
		FileOutputStream file = new FileOutputStream(new File("NotasSOO.xlsx"));

		// Criando pasta de trabalho
		Workbook workbook = new HSSFWorkbook();

		// Criando a planilha notas
		Sheet planilha = workbook.createSheet("Notas");

		// Definindo alguns padrões
		planilha.setDefaultColumnWidth(15);
		planilha.setDefaultRowHeight((short) 400);

		// Criando linha 0 e celulas
		Row linha = planilha.createRow(0);
		Cell celula = linha.createCell(0);

		// Fonte
		Font font = workbook.createFont();
		font = getFontHeader(font);

		// Estilo da celula
		CellStyle style = workbook.createCellStyle();
		getStyleHeader(style, linha, celula, font, planilha);

		// Setando valores
		writeSheet(linha, celula, "Wesley", 9, 1, planilha);
		writeSheet(linha, celula, "Bruno", 7, 2, planilha);
		writeSheet(linha, celula, "Bianca", 10, 3, planilha);
		writeSheet(linha, celula, "Rafael", 6, 4, planilha);
		writeSheet(linha, celula, "Sumara", 9, 5, planilha);
		writeSheet(linha, celula, "Nicole", 7, 6, planilha);
		writeSheet(linha, celula, "Guilherm", 10, 7, planilha);
		writeSheet(linha, celula, "Rafael Moreira", 6, 8, planilha);
		writeSheet(linha, celula, "Ricado ", 6, 9, planilha);
		
			
		
		// Funcao Media
		mathFunction(celula, linha, planilha);

		// Passar os dados que foram inseridos para o arquivo
		workbook.write(file);

		// Fechamento do arquivo
		file.close();

		// Arquivo foi criado com sucesso
		JOptionPane.showMessageDialog(null, "Arquivo Criado Com sucesso");

	}

	// Método de fonte do cabeçalho

	public static Font getFontHeader(Font font) {

		font.setFontHeightInPoints((short) 10);
		font.setFontName("Arial");
		font.setColor(IndexedColors.WHITE.getIndex());
		font.setBold(true);
		font.setItalic(false);

		return font;
	}

	// Crio meu cabeçalho e aplico tanto estilo/fonte

	public static void getStyleHeader(CellStyle style, Row linha, Cell celula, Font font, Sheet planilha) {

		style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		celula = linha.createCell(0);
		celula.setCellValue(new HSSFRichTextString("Nome"));
		style.setFont(font);
		celula.setCellStyle(style);
		
		celula = linha.createCell(1);
		celula.setCellValue(new HSSFRichTextString("Nota"));
		style.setFont(font);
		celula.setCellStyle(style);
	
	}

	// Método para escrever na planilha
	public static void writeSheet(Row linha, Cell celula, String nome, int nota, int numeroLinha, Sheet planilha) {

		linha = planilha.createRow(numeroLinha);
		celula = linha.createCell(0);
		celula.setCellValue(nome);
		celula = linha.createCell(1);
		celula.setCellValue(nota);

	}

	// Metodo Funcao metematica

	public static void mathFunction(Cell celula, Row linha, Sheet planilha) {

		linha = planilha.createRow(10);
		celula = linha.createCell(1);
		celula.setCellFormula("AVERAGE(B2:B10)");
	 	
	}
	

}
