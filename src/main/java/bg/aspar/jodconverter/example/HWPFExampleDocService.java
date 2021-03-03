package bg.aspar.jodconverter.example;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hwpf.HWPFDocument;

import bg.aspar.jodconverter.service.VariableReplacerForDoc;
import bg.aspar.jodconverter.service.impl.VariableReplacerForDocImpl;

public class HWPFExampleDocService {

	public static void main(String[] args) throws IOException {
		String inputFileName = "Example-01-template.doc";
		String resultFileName = "Example-01-result.doc";

		FileInputStream fileInStream = new FileInputStream(inputFileName);

		VariableReplacerForDoc docConverter = new VariableReplacerForDocImpl();
		HWPFDocument doc = docConverter.convert(fileInStream);

		saveWord(resultFileName, doc);
	}

	private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException {
		try (FileOutputStream out = new FileOutputStream(filePath)) {
			doc.write(out);
		}
	}

}
