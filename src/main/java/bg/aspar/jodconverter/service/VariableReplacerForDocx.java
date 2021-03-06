package bg.aspar.jodconverter.service;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;


public interface VariableReplacerForDocx {

	public XWPFDocument convert(InputStream inStream) throws IOException;

}
