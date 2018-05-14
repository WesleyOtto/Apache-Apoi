import java.io.File;
import java.io.FileOutputStream;
import javax.swing.JOptionPane;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class Principal {

	public static void main(String[] args) throws Exception {

		// Cria o documento em branco
		XWPFDocument document = new XWPFDocument();

		// Criando o arquivo de escrita
		FileOutputStream file = new FileOutputStream(new File("Doc.docx"));

		// Titulo
		XWPFParagraph t1 = document.createParagraph(); // Cria um paragrafo em branco
		t1.setAlignment(ParagraphAlignment.CENTER); // Alinho o meu pargrafo no centro
		XWPFRun titulo = t1.createRun(); // Crio o titulo do meu texto
		styleRun(titulo, "Star Wars", "Arial", "FF0000", 18);

		// TEXTO
		XWPFParagraph p2 = document.createParagraph(); // Crio o Paragrafo para o corpo do texto
		p2.setAlignment(ParagraphAlignment.BOTH); // Alinho o paragrafo para justificar

		XWPFRun body = p2.createRun(); // Crio o meu corpo do texto
		body.setFontFamily("Arial");
		body.setFontSize(12);
		body.addTab(); // Tab para iniciar o texto
		body.setText(
				"Star Wars (Guerra nas Estrelas (t�tulo no Brasil) ou Guerra das Estrelas (t�tulo em Portugal)) � uma franquia do tipo "
						+ "space opera estadunidense criada pelo cineasta George Lucas que conta com uma s�rie de oito filmes de fantasia cient�fica e um spin-off. "
						+ "O primeiro filme foi lan�ado apenas com o t�tulo Star Wars em 25 de maio de 1977"
						+ ", e tornou-se um fen�meno mundial inesperado de cultura popular, sendo respons�vel pelo in�cio da era dos blockbusters: "
						+ "Super produ��es cinematogr�ficas que fazem sucesso nas bilheterias e viram franquias com brinquedos, jogos, livros, etc. Foi "
						+ "seguido por duas sequ�ncias, The Empire Strikes Back e Return of the Jedi, lan�adas com intervalos de tr�s anos. Esta primeira "
						+ "trilogia segue o trio ic�nico: Luke Skywalker, Han Solo e Princesa Leia, que luta na Alian�a Rebelde para derrubar o tirano Imp�rio "
						+ "Gal�ctico; paralelamente ocorre a jornada de Luke para se tornar um cavaleiro Jedi e a luta contra Darth Vader, um ex-Jedi que sucumbiu ao Lado Sombrio da For�a e ao Imperador.");

		body.addBreak();
		body.addBreak();
		body.addBreak();

		// Titulo
		XWPFParagraph r1 = document.createParagraph();
		r1.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun referencias = r1.createRun();
		styleRun(referencias, "Referencias Bibliogr�ficas", "Arial", "238E68", 18);

		// Referencias
		XWPFParagraph WebSite = document.createParagraph();
		XWPFRun ref = WebSite.createRun(); // Crio o meu corpo do texto
		ref.setText("__________. STAR WARS. Dispon�vel em: <https://pt.wikipedia.org/wiki/Star_Wars>. Acessado em 13 de maio de 2018 �s 23h16min. ");

		document.write(file);
		file.close();
		document.close();

		// Arquivo foi criado com sucesso
		JOptionPane.showMessageDialog(null, "Arquivo Criado Com sucesso");
	}

	public static void styleRun(XWPFRun run, String text, String fonte, String color, int tamanho) {

		run.setBold(true); // Negrito
		run.setItalic(true); // Italico
		run.setText(text); // Seto o texto
		run.setFontFamily(fonte); // TIpo da fonte
		run.setFontSize(tamanho); // Tamanho da fonte
		run.setColor(color); // Seto a cor
		run.addBreak(); // Pula linha

	}
}
