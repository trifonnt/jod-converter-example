package bg.aspar.jodconverter.example;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;

import bg.aspar.jodconverter.service.VariableReplacerForDoc;
import bg.aspar.jodconverter.service.impl.VariableReplacerForDocImpl;

public class HWPFExampleDocService {

	public static void main(String[] args) throws IOException {
		String inputFileName = "Example-01-template.doc";
		String resultFileName = "Example-01-result.doc";

		Map<String, String> variables = new HashMap<>();
		variables.put("${var01}", "MyValue1");
		variables.put("${var02}", "MyValue2");
		variables.put("${var03}", "MyValue3");

		FileInputStream fileInStream = new FileInputStream(inputFileName);

		VariableReplacerForDoc docConverter = new VariableReplacerForDocImpl();
		HWPFDocument doc = docConverter.convert(fileInStream, variables);

		saveWord(resultFileName, doc);
	}

	private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException {
		try (FileOutputStream out = new FileOutputStream(filePath)) {
			doc.write(out);
		}
	}

}
