package it.maior.icaro.word.docx;

import java.math.BigInteger;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.P;
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.PPrBase.NumPr.Ilvl;
import org.docx4j.wml.PPrBase.NumPr.NumId;

/**
 * Creates a WordprocessingML document from scratch,
 * including a numbering definitions part, and use
 * it to demonstrate restart numbering.
 *
 * @author Jason Harrop
 */
public class NumberingRestart {

	static org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();

	static String filename = "OUT_NumberingRestart.docx";

	public static void main(String[] args) throws Exception {

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

		// Add numbering part
		NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
		wordMLPackage.getMainDocumentPart().addTargetPart(ndp);
		ndp.setJaxbElement( (Numbering) XmlUtils.unmarshalString(initialNumbering) );

		// Add some document content
		wordMLPackage.getMainDocumentPart().addParagraphOfText("Example of restarting numbering");

		P p = createBulletParagraph(1, 0, "text on top level" );
		wordMLPackage.getMainDocumentPart().addObject(p);

		wordMLPackage.getMainDocumentPart().addObject(
				createBulletParagraph(1, 0, "more text on top level" ));

		wordMLPackage.getMainDocumentPart().addObject(
				createBulletParagraph(1, 1, "text on level 1" ));

		// Ok, lets restart numbering
		long newNumId = ndp.restart(1, 0, 1);

		wordMLPackage.getMainDocumentPart().addObject(
				createBulletParagraph(newNumId, 0, "text on top level - restarted" ));

		// After first using newNumId, it doesn't matter whether
		// subsequent paragraphs use that or the original numId
		wordMLPackage.getMainDocumentPart().addObject(
				createBulletParagraph(newNumId, 0, "text on top level - using newNumId" ));

		wordMLPackage.getMainDocumentPart().addObject(
				createBulletParagraph(1, 0, "text on top level - using original NumId" ));

		// Now save it
		wordMLPackage.save(new java.io.File(System.getProperty("user.dir") + "/" + filename) );

		System.out.println("Done. Saved " + filename);

	}

	/**
	 * Create something like:
	 *
	 *               <w:p>
	 <w:pPr>
	 <w:numPr>
	 <w:ilvl w:val="0"/>
	 <w:numId w:val="1"/>
	 </w:numPr>
	 </w:pPr>
	 <w:r>
	 <w:t>B</w:t>
	 </w:r>
	 </w:p>

	 * @return
	 */
	private static P createBulletParagraph(long numId, long ilvl, String paragraphText ) {

		P  p = factory.createP();

		org.docx4j.wml.Text  t = factory.createText();
		t.setValue(paragraphText);

		org.docx4j.wml.R  run = factory.createR();
		run.getContent().add(t);

		p.getContent().add(run);

		org.docx4j.wml.PPr ppr = factory.createPPr();
		p.setPPr( ppr );

		// Create and add <w:numPr>
		NumPr numPr =  factory.createPPrBaseNumPr();
		ppr.setNumPr(numPr);

		// The <w:ilvl> element
		Ilvl ilvlElement = factory.createPPrBaseNumPrIlvl();
		numPr.setIlvl(ilvlElement);
		ilvlElement.setVal(BigInteger.valueOf(ilvl));

		// The <w:numId> element
		NumId numIdElement = factory.createPPrBaseNumPrNumId();
		numPr.setNumId(numIdElement);
		numIdElement.setVal(BigInteger.valueOf(numId));

		return p;

	}


	static final String initialNumbering = "<w:numbering xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"
			+ "<w:abstractNum w:abstractNumId=\"0\">"
			+ "<w:nsid w:val=\"2DD860C0\"/>"
			+ "<w:multiLevelType w:val=\"multilevel\"/>"
			+ "<w:tmpl w:val=\"0409001D\"/>"
			+ "<w:lvl w:ilvl=\"0\">"
			+ "<w:start w:val=\"1\"/>"
			+ "<w:numFmt w:val=\"bullet\"/>"
			+ "<w:lvlText w:val=\"•\"/>"
			+ "<w:lvlJc w:val=\"left\"/>"
			+ "<w:pPr>"
			+ "<w:ind w:left=\"360\" w:hanging=\"360\"/>"
			+ "</w:pPr>"
			+ "</w:lvl>"
			+ "<w:lvl w:ilvl=\"1\">"
			+ "<w:start w:val=\"1\"/>"
			+ "<w:numFmt w:val=\"bullet\"/>"
			+ "<w:lvlText w:val=\"○\"/>"
			+ "<w:lvlJc w:val=\"left\"/>"
			+ "<w:pPr>"
			+ "<w:ind w:left=\"720\" w:hanging=\"360\"/>"
			+ "</w:pPr>"
			+ "</w:lvl>"
			+ "<w:lvl w:ilvl=\"2\">"
			+ "<w:start w:val=\"1\"/>"
			+ "<w:numFmt w:val=\"bullet\"/>"
			+ "<w:lvlText w:val=\"◘\"/>"
			+ "<w:lvlJc w:val=\"left\"/>"
			+ "<w:pPr>"
			+ "<w:ind w:left=\"1080\" w:hanging=\"360\"/>"
			+ "</w:pPr>"
			+ "</w:lvl>"
			+ "</w:abstractNum>"
			+ "<w:num w:numId=\"1\">"
			+ "<w:abstractNumId w:val=\"0\"/>"
			+ "</w:num>"
			+ "</w:numbering>";
}
