package bg.aspar.jodconverter.service.impl;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Service;

import bg.aspar.jodconverter.service.VariableReplacerForDocx;

@Service
public class VariableReplacerForDocxImpl implements VariableReplacerForDocx {

	@Override
	public XWPFDocument convert(InputStream inStream) throws IOException {
		XWPFDocument doc = new XWPFDocument(inStream);

		doc = replaceText(doc, "${var01}", "MyValue1");

		return doc;
	}

	private static XWPFDocument replaceText(XWPFDocument doc, String findText, String replaceText) {
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
/*
	private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException {
		try (FileOutputStream out = new FileOutputStream(filePath)) {
			doc.write(out);
		}
	}
*/
}
