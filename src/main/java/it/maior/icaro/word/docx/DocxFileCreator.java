package it.maior.icaro.word.docx;

import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.*;
import org.docx4j.wml.*;
import org.docx4j.wml.Color;
import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Files;
import java.util.List;

public class DocxFileCreator {

    public WordprocessingMLPackage wordPackage;
    public MainDocumentPart mainDocumentPart;

    public DocxFileCreator() {
        init();
    }

    private void init() {
        try {
            this.wordPackage = WordprocessingMLPackage.createPackage();
            this.mainDocumentPart= this.wordPackage.getMainDocumentPart();

            // Add numbering part
            NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
            this.wordPackage.getMainDocumentPart().addTargetPart(ndp);
            ndp.setJaxbElement( (Numbering) XmlUtils.unmarshalString(initialNumbering) );

        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (JAXBException e) {
            e.printStackTrace();
        }
    }

    public  void addStyledParagraphOfText (String styleId, String text){
        mainDocumentPart.addStyledParagraphOfText(styleId, text);
    }

    public  void addParagraphOfText (String text){
        mainDocumentPart.addParagraphOfText(text);
    }

    // MATO: it creates a paragraph with a hard coded configuration: object r and rpr
    public void addParagraphWithConfiguration(String s) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P p = factory.createP();
        R r = factory.createR();
        Text t = factory.createText();
        t.setValue(s);
        r.getContent().add(t);
        p.getContent().add(r);
        RPr rpr = factory.createRPr();
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        rpr.setB(b);
        rpr.setI(b);
        rpr.setCaps(b);
        Color red = factory.createColor();
        red.setVal("green");
        rpr.setColor(red);
        r.setRPr(rpr);
        mainDocumentPart.getContent().add(p);
    }

    //MATO: it creates a table with dimension equal to rowNumber X columnNumber where every cell contains the same value (value parameter)
    public void addTable(int columnNumber, int rowNumber, String value){
        int writableWidthTwips = wordPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
        Tbl tbl = TblFactory.createTable(columnNumber, rowNumber, writableWidthTwips / columnNumber);
        List<Object> rows = tbl.getContent();
        for (Object row : rows) {
            Tr tr = (Tr) row;
            List<Object> cells = tr.getContent();
            for (Object cell : cells) {
                Tc td = (Tc) cell;
                td.getContent().add(value);
            }
        }
    }

    public void addImage(String imagePath, String filenameHint){

        try {
            File image = new File(imagePath);
            byte[] fileContent = Files.readAllBytes(image.toPath());
            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordPackage, fileContent);
            Inline inline = imagePart.createImageInline(filenameHint, "Alt Text", 1, 2, false);
            P Imageparagraph = addImageToParagraph(inline);
            mainDocumentPart.getContent().add(Imageparagraph);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private P addImageToParagraph(Inline inline) {
        ObjectFactory factory = new ObjectFactory();
        P p = factory.createP();
        R r = factory.createR();
        p.getContent().add(r);
        Drawing drawing = factory.createDrawing();
        r.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        return p;
    }


    public void saveFileDocx(String outputPath) {
        try {
            File exportFile = new File(outputPath);
            wordPackage.save(exportFile);
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
    }


    public void addBulletParagraphOfText(long numId, long ilvl, String paragraphText ) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P  p = factory.createP();

        org.docx4j.wml.Text  t = factory.createText();
        t.setValue(paragraphText);

        org.docx4j.wml.R  run = factory.createR();
        run.getContent().add(t);

        p.getContent().add(run);

        org.docx4j.wml.PPr ppr = factory.createPPr();
        p.setPPr( ppr );

        // Create and add <w:numPr>
        PPrBase.NumPr numPr =  factory.createPPrBaseNumPr();
        ppr.setNumPr(numPr);

        // The <w:ilvl> element
        PPrBase.NumPr.Ilvl ilvlElement = factory.createPPrBaseNumPrIlvl();
        numPr.setIlvl(ilvlElement);
        ilvlElement.setVal(BigInteger.valueOf(ilvl));

        // The <w:numId> element
        PPrBase.NumPr.NumId numIdElement = factory.createPPrBaseNumPrNumId();
        numPr.setNumId(numIdElement);
        numIdElement.setVal(BigInteger.valueOf(numId));

        this.wordPackage.getMainDocumentPart().addObject(p);
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
