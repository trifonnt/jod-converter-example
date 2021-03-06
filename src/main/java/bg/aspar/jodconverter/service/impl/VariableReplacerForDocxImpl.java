package bg.aspar.jodconverter.service.impl;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Service;

import bg.aspar.jodconverter.service.VariableReplacerForDocx;

@Service
public class VariableReplacerForDocxImpl implements VariableReplacerForDocx {

	@Override
	public XWPFDocument convert(InputStream inStream, Map<String, String> variables) throws IOException {
		XWPFDocument doc = new XWPFDocument(inStream);

		doc = replaceText(doc, variables);

		return doc;
	}

	private static XWPFDocument replaceText(XWPFDocument doc, Map<String, String> variables) {
		doc.getParagraphs().forEach(p -> {
			p.getRuns().forEach(run -> {
				String text = run.text();
				for (Map.Entry<String, String> entry : variables.entrySet()) {
					if (text.contains(entry.getKey())) {
						text = text.replace(entry.getKey(), entry.getValue());
						run.setText(text, 0);
					}
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
