import com.aspose.words.*;

import java.util.regex.Pattern;


public class Main {

    public static void main(String[] args) throws Exception {
        // write your code here
        System.out.println("Hello World");

        //load files
        DocumentLoader documentLoader = new DocumentLoader();

        Document outerDoc = documentLoader.getDocument("outerDoc.docx");
        Document innerDoc1 = documentLoader.getDocument("innerDoc1.docx");
        Document innerDoc2 = documentLoader.getDocument("innerDoc2.docx");

        Document partialDoc = insertDocumentAtReplace(outerDoc, innerDoc1, "innerDoc1_placeholder");
        Document finalDoc = insertDocumentAtReplace(partialDoc, innerDoc2, "innerDoc2_placeholder");
        
        finalDoc.save("finalOuterDoc.docx");
    }

    public static Document insertDocumentAtReplace(Document mainDoc, Document innerDoc, String placeholder) throws Exception {
        InsertDocumentAtReplaceHandler findDocumentSnippetsHandler = new InsertDocumentAtReplaceHandler(innerDoc);
        FindReplaceOptions replacementOptions = new FindReplaceOptions();
        replacementOptions.setReplacingCallback(findDocumentSnippetsHandler);
        mainDoc.getRange().replace(Pattern.compile("\\{"+placeholder+"\\}"), "", replacementOptions);
        return mainDoc;
    }

    private static class InsertDocumentAtReplaceHandler implements IReplacingCallback {

        private Document innerDocument;

        private InsertDocumentAtReplaceHandler(Document innerDocument) {
            this.innerDocument = innerDocument;
        }

        public int replacing(ReplacingArgs e) throws Exception {
            // Insert a document after the paragraph, containing the match text.
            Paragraph para = (Paragraph) e.getMatchNode().getParentNode();
            insertDocument(para, innerDocument);

            // Remove the paragraph with the match text.
            para.remove();
            return ReplaceAction.SKIP;
        }
    }

    public static void insertDocument(Node insertAfterNode, Document srcDoc) throws Exception {
        // Make sure that the node is either a paragraph or table.
        if ((insertAfterNode.getNodeType() != NodeType.PARAGRAPH) & (insertAfterNode.getNodeType() != NodeType.TABLE))
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");

        // We will be inserting into the parent of the destination paragraph.
        CompositeNode dstStory = insertAfterNode.getParentNode();

        // This object will be translating styles and lists during the import.
        NodeImporter importer = new NodeImporter(srcDoc, insertAfterNode.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // Loop through all sections in the source document.
        for (Section srcSection : srcDoc.getSections()) {
            // Loop through all block level nodes (paragraphs and tables) in the body of the section.
            for (Node srcNode : (Iterable<Node>) srcSection.getBody()) {
                // Let's skip the node if it is a last empty paragraph in a section.
                if (srcNode.getNodeType() == (NodeType.PARAGRAPH)) {
                    Paragraph para = (Paragraph) srcNode;
                    if (para.isEndOfSection() && !para.hasChildNodes())
                        continue;
                }

                // This creates a clone of the node, suitable for insertion into the destination document.
                Node newNode = importer.importNode(srcNode, true);

                // Insert new node after the reference node.
                dstStory.insertAfter(newNode, insertAfterNode);
                insertAfterNode = newNode;
            }
        }
    }
}
