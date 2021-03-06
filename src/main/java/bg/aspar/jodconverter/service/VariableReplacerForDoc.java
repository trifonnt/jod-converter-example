package bg.aspar.jodconverter.service;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;

public interface VariableReplacerForDoc {

	public HWPFDocument convert(InputStream inStream, Map<String, String> variables) throws IOException;

}
