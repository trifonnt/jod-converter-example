package bg.aspar.jodconverter.example;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class HWPFExampleDoc {

	public static void main(String[] args) {
		String inputFileName = "Example-01-template.doc";
		String resultFileName = "Example-01-result.doc";

		try (POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(inputFileName))) {
//			fs = new POIFSFileSystem(new FileInputStream(filePath));
			HWPFDocument doc = null;
			try {
				doc = new HWPFDocument(fs);

				doc = replaceText(doc, "${var01}", "MyValue1");

				saveWord(resultFileName, doc);
			} finally {
				if (doc != null) {
					doc.close();
				}
			}
		} catch (IOException ex) {
			ex.printStackTrace();
		}
	}

	private static HWPFDocument replaceText(HWPFDocument doc, String findText, String replaceText) {
		Range r1 = doc.getRange();

		for (int i = 0; i < r1.numSections(); ++i) {
			Section s = r1.getSection(i);
			for (int x = 0; x < s.numParagraphs(); x++) {
				Paragraph p = s.getParagraph(x);
				for (int z = 0; z < p.numCharacterRuns(); z++) {
					CharacterRun run = p.getCharacterRun(z);
					String text = run.text();
					if (text.contains(findText)) {
						run.replaceText(findText, replaceText);
					}
				}
			}
		}
		return doc;
	}

	private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException {
		try (FileOutputStream out = new FileOutputStream(filePath)) {
			doc.write(out);
		}
	}
}
