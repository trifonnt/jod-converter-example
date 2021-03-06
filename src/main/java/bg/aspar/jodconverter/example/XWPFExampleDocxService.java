package bg.aspar.jodconverter.example;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import bg.aspar.jodconverter.service.VariableReplacerForDocx;
import bg.aspar.jodconverter.service.impl.VariableReplacerForDocxImpl;

public class XWPFExampleDocxService {

	public static void main(String[] args) throws IOException {
		String inputFileName = "Example-01-template.docx";
		String resultFileName = "Example-01-result.docx";

		Map<String, String> variables = new HashMap<>();
		variables.put("${var01}", "MyValue1");
		variables.put("${var02}", "MyValue2");
		variables.put("${var03}", "MyValue3");

		FileInputStream fileInStream = new FileInputStream(inputFileName);

		VariableReplacerForDocx variableReplacer = new VariableReplacerForDocxImpl();
		XWPFDocument doc = variableReplacer.convert(fileInStream, variables);

		saveWord(resultFileName, doc);
	}

	private static void saveWord(String filePath, XWPFDocument doc) throws FileNotFoundException, IOException {
		try (FileOutputStream out = new FileOutputStream(filePath)) {
			doc.write(out);
		}
	}

}
