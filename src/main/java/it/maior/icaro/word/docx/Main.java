package it.maior.icaro.word.docx;


public class Main {

    static String outputPath = "./OutputFiles/example.docx";

    public static void main(String[] args) {

        DocxFileCreator dfc = new DocxFileCreator();

        dfc.addStyledParagraphOfText("Title", "IcaroXt: From Jira to Docx");
        dfc.addStyledParagraphOfText("Subtitle", "An experiment");

        dfc.addStyledParagraphOfText("Heading1", "Add an image");
        dfc.addImage("C:/Users/mato/Pictures/Capture.PNG", "Hint");

        dfc.addStyledParagraphOfText("Heading1", "Paragraph");
        dfc.addStyledParagraphOfText("Heading2", "Paragraph simple format");
        dfc.addParagraphOfText("This element specifies the set of run properties applied to the glyph used to represent the physical location of the paragraph mark for this paragraph. This paragraph mark, being a physical character in the document, can be formatted, and therefore shall be capable of representing this formatting like any other character in the document. If this element is not present, the paragraph mark is unformatted, as with any other run of text.");
        dfc.addStyledParagraphOfText("Heading2", "Paragraph with special format");
        dfc.addParagraphWithConfiguration("This element specifies the set of run properties applied to the glyph used to represent the physical location of the paragraph mark for this paragraph. This paragraph mark, being a physical character in the document, can be formatted, and therefore shall be capable of representing this formatting like any other character in the document. If this element is not present, the paragraph mark is unformatted, as with any other run of text.");

        dfc.addStyledParagraphOfText("Heading1", "Bullet list");
        dfc.addParagraphOfText("This element specifies the set of run properties:");
        dfc.addBulletParagraphOfText(1, 0, "text on top level");
        dfc.addBulletParagraphOfText(1, 0, "more text on top level");
        dfc.addBulletParagraphOfText(1, 1, "text on level 1");

        dfc.saveFileDocx(outputPath);
    }
}