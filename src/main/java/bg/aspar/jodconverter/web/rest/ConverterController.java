package bg.aspar.jodconverter.web.rest;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.util.Assert;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.document.DocumentFormat;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.util.FileUtils;
import org.jodconverter.core.util.StringUtils;

@Controller
public class ConverterController {

	private static final String ATTRNAME_ERROR_MESSAGE = "errorMessage";
	private static final String ON_ERROR_REDIRECT = "redirect:/";

	@Autowired
	private DocumentConverter converter;

	@GetMapping("/")
	/* default */ String index() {
		return "converter";
	}

	/**
	 * Converts a source file to a target format.
	 *
	 * @param inputFile          Source file to convert.
	 * @param outputFormat       Output format of the conversion.
	 * @param redirectAttributes Model that contains attributes
	 * @return The converted file, or the error redirection if an error occurs.
	 */
	@PostMapping("/converter")
	/* default */ Object convert(
			@RequestParam(name = "inputFile") final MultipartFile inputFile,
			@RequestParam(name = "outputFormat", required = false) final String outputFormat,
			final RedirectAttributes redirectAttributes) 
	{

		if (inputFile.isEmpty()) {
			redirectAttributes.addFlashAttribute(ATTRNAME_ERROR_MESSAGE, "Please select a file to upload.");
			return ON_ERROR_REDIRECT;
		}

		if (StringUtils.isBlank(outputFormat)) {
			redirectAttributes.addFlashAttribute(ATTRNAME_ERROR_MESSAGE, "Please select an output format.");
			return ON_ERROR_REDIRECT;
		}

		// Here, we could have a dedicated service that would convert document
		try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

			final DocumentFormat targetFormat = DefaultDocumentFormatRegistry.getFormatByExtension(outputFormat);
			Assert.notNull(targetFormat, "targetFormat must not be null");
			converter.convert(inputFile.getInputStream())
				.to(baos)
				.as(targetFormat)
				.execute();

			final HttpHeaders headers = new HttpHeaders();
			headers.setContentType(MediaType.parseMediaType(targetFormat.getMediaType()));
			headers.add("Content-Disposition", "attachment; filename="
					+ FileUtils.getBaseName(inputFile.getOriginalFilename()) + "." + targetFormat.getExtension());
			return new ResponseEntity<>(baos.toByteArray(), headers, HttpStatus.OK);

		} catch (OfficeException | IOException ex) {
			redirectAttributes.addFlashAttribute(ATTRNAME_ERROR_MESSAGE,
					"Could not convert the file " + inputFile.getOriginalFilename() + ". Cause: " + ex.getMessage());
		}

		return ON_ERROR_REDIRECT;
	}

	
}