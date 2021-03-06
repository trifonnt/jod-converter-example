package bg.aspar.jodconverter.service.impl;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.springframework.stereotype.Service;

import bg.aspar.jodconverter.service.VariableReplacerForDoc;

@Service
public class VariableReplacerForDocImpl implements VariableReplacerForDoc {

	@Override
	public HWPFDocument convert(InputStream inStream, Map<String, String> variables) throws IOException {
		HWPFDocument doc = new HWPFDocument(inStream);

		doc = replaceText(doc, variables);

		return doc;
	}

	private static HWPFDocument replaceText(HWPFDocument doc, Map<String, String> variables) {
		Range r1 = doc.getRange();

		for (int i = 0; i < r1.numSections(); ++i) {
			Section s = r1.getSection(i);
			for (int x = 0; x < s.numParagraphs(); x++) {
				Paragraph p = s.getParagraph(x);
				for (int z = 0; z < p.numCharacterRuns(); z++) {
					CharacterRun run = p.getCharacterRun(z);
					String text = run.text();
					for (Map.Entry<String, String> variable : variables.entrySet()) {
						if (text.contains(variable.getKey())) {
							run.replaceText(variable.getKey(), variable.getValue());
						}
					}
				}
			}
		}
		return doc;
	}
/*
	private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException {
		try (FileOutputStream out = new FileOutputStream(filePath)) {
			doc.write(out);
		}
	}
*/
}
