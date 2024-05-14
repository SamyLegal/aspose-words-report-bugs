package fr.mgdis.aspose.words;

import com.aspose.words.Document;
import com.aspose.words.IMailMergeDataSourceRoot;
import com.aspose.words.LayoutCollector;
import com.aspose.words.LoadOptions;
import com.aspose.words.MailMergeCleanupOptions;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.Section;
import io.quarkus.logging.Log;
import jakarta.enterprise.context.ApplicationScoped;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import org.apache.commons.lang3.StringUtils;

@ApplicationScoped
public class AsposeWordsBuilder {
  /**
   * Function that remove empty sections and paragraphs
   *
   * @param doc Word document to clean
   */
  private void removeEmptySectionsAndParagraphs(Document doc) {
    // Remove empty sections if some persists in the document
    // https://forum.aspose.com/t/remove-blank-pages-from-word-document-with-headers-footers-using-c-net-or-java/222114/8
    while (doc.getSections().getCount() > 0
      && doc.getLastSection().getBody().getChildNodes(NodeType.RUN, true).getCount() == 0
      && doc.getLastSection().getBody().getChildNodes(NodeType.SHAPE, true).getCount() == 0) {
      doc.getLastSection().remove();
    }

    // Remove empty paragraphs at the end of each section
    for (Section section : doc.getSections()) {
      while (section.getBody().getLastChild().getNodeType() == NodeType.PARAGRAPH
        && !section.getBody().getLastParagraph().hasChildNodes()) {
        section.getBody().getLastParagraph().remove();
      }
    }
  }

  /**
   * Function that return the number of blank pages contains in the document
   *
   * @param doc Document
   * @return List of blank pages
   * @throws Exception Exception
   */
  /* private java.util.List<Integer> getNumberOfBlankPages(Document doc) throws Exception {
    // Page numbers of blank pages
    List<Integer> blankPages = new ArrayList<>();

    // Scan the document by page to remove blank pages
    for (int page = 0; page < doc.getPageCount(); page++) {

      // Extract each page separately to analyze its content
      Document pageDoc = doc.extractPages(page, 1);

      // Get text content of this page
      StringBuilder textOfPage = new StringBuilder();
      for (Section section : pageDoc.getSections()) {
        // Let's not consider the content of Headers and Footers
        textOfPage.append(section.getBody().toString(SaveFormat.TEXT));
        if (StringUtils.isBlank(textOfPage.toString().trim())) {
          blankPages.add(page);
        }
      }
    }

    return blankPages;
  } */

  /**
   * Function that removes blank pages from document.
   * The cause of the creation of blank pages is the offset of a page break in the flow of the Word document.
   * This behavior will happen randomly depending on the dataset sent to the document.
   * <a href="https://forum.aspose.com/t/code-to-remove-empty-pages-from-word-document-using-c-net/214034/9">Post forum that explain the solution</a>
   *
   * @param doc Document to analyze
   */
  /* private void removeBlankPages(Document doc, List<Integer> blankPages) throws Exception {
    // LayoutCollector is a class that allows you to transform the flow of a Word document into a document with a fixed page.
    LayoutCollector layoutCollector = new LayoutCollector(doc);

    List<Node> list = new ArrayList<>();
    for (Node node : doc.getChildNodes(NodeType.ANY, true).toArray()) {
      if (layoutCollector.getNumPagesSpanned(node) == 0) {
        int pageIndex = layoutCollector.getStartPageIndex(node);
        if (blankPages.contains(pageIndex - 1)) {
          list.add(node);
        }
      }
    }

    for (Node node : list) {
      node.remove();
    }
  } */

  /**
   * Build document with Aspose Words
   *
   * @param templateInput  Template input.
   * @param dataSourceRoot Data source
   * @return Document
   * @throws Exception Exception
   */
  public ByteArrayOutputStream buildDocument(InputStream templateInput, IMailMergeDataSourceRoot dataSourceRoot)
    throws Exception {
    LoadOptions loadOptions = new LoadOptions();
    Document doc = new Document(templateInput, loadOptions);

    var mailMerge = doc.getMailMerge();

    // REMOVE_UNUSED_REGIONS : Option that remove section in the template that use data sources that does not exist
    // REMOVE_EMPTY_PARAGRAPHS : Option that remove empty lines in the generated document when data sources exist but is empty.
    // The deleted lines will only be those containing empty "merge fields", ":" and spaces.
    // If an alphanumeric character is present the paragraph will be generated.
    mailMerge.setCleanupOptions(
      MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS | MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS |
        MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS);

    // Add callback to enable html text in merge field
    // (http://www.aspose.com/community/forums/thread/380671/html-text-with-merge-field.aspx)
    mailMerge.setMergeDuplicateRegions(true);
    // mailMerge.setRestartListsAtEachSection(true);
    mailMerge.setFieldMergingCallback(new HandleMergeFields());

    // Authorize mustache syntax in word template
    mailMerge.setUseNonMergeFields(true);

    // Authorize to use the same region with same name multiple times
    mailMerge.setMergeDuplicateRegions(true);
    mailMerge.executeWithRegions(dataSourceRoot);

    // Fusion object is a global object that contains all data sources.
    // Send "fusion" object authorize to use data sources directly
    // with simple syntaxe like this "MERGEFIELD dossierFinancement.reference" or much simpler "{{ dossierFinancement.reference }}"
    mailMerge.execute(dataSourceRoot.getDataSource("fusion"));

    // Close template stream
    templateInput.close();

    // Removing unused sections and empty paragraphs for have a clean document
    // We calculate blank pages after this clean for remove some blank pages persistent after this clean
    removeEmptySectionsAndParagraphs(doc);

    // Check if document has blank pages
    // var blankPages = getNumberOfBlankPages(doc);
    // Log.debug("[AsposeWordsBuilder] - Is that document contains blank pages : " + !blankPages.isEmpty());

    // Remove blank pages of our document
    /* if (!blankPages.isEmpty()) {
      var numberOfBlankPages = blankPages.stream().map(Object::toString).collect(Collectors.joining(","));
      Log.debug("[AsposeWordsBuilder] - Blank pages number in the document : " + numberOfBlankPages);
      removeBlankPages(doc, blankPages);
      Log.debug("[AsposeWordsBuilder] - Blank pages removed from the document");
    } */

    doc.updateFields();

    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    doc.save(outputStream, new PdfSaveOptions());

    return outputStream;
  }
}
