package bg.aspar.jodconverter.service;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;


public interface VariableReplacerForDocx {

	public XWPFDocument convert(InputStream inStream, Map<String, String> variables) throws IOException;

}
