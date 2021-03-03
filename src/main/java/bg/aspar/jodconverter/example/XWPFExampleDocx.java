package bg.aspar.jodconverter.example;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class XWPFExampleDocx {

	public static void main(String[] args) throws URISyntaxException, IOException {
		String inputFileName = "Example-01-template.docx";
		String resultFileName = "Example-01-result.docx";

//		Path templatePath = Paths.get(XWPFExampleDocx.class.getClassLoader().getResource(inputFileName).toURI());
		Path templatePath = Paths.get(inputFileName);
	
		XWPFDocument doc = new XWPFDocument(Files.newInputStream(templatePath));
		doc = replaceTextFor(doc, "${var01}", "MyValue1");

		saveWord(resultFileName, doc);
	}

	private static XWPFDocument replaceTextFor(XWPFDocument doc, String findText, String replaceText) {
		doc.getParagraphs().forEach(p -> {
			p.getRuns().forEach(run -> {
				String text = run.text();
				if (text.contains(findText)) {
					run.setText(text.replace(findText, replaceText), 0);
				}
			});
		});

		return doc;
	}

	private static void saveWord(String filePath, XWPFDocument doc) throws FileNotFoundException, IOException {
		try (FileOutputStream out = new FileOutputStream(filePath)) {
			doc.write(out);
		}
	}
}
