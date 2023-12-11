package fr.mgdis.aspose.words;

import com.aspose.words.*;
import jakarta.enterprise.context.ApplicationScoped;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

@ApplicationScoped
public class AsposeWordsBuilder {
  /**
   * Build document with Aspose Words
   *
   * @param templateInput  Template input.
   * @param dataSourceRoot Data source
   * @return Document
   * @throws Exception Exception
   */
  public ByteArrayOutputStream buildDocument(InputStream templateInput, IMailMergeDataSourceRoot dataSourceRoot) throws Exception {
    LoadOptions loadOptions = new LoadOptions();
    Document doc = new Document(templateInput, loadOptions);

    var mailMerge = doc.getMailMerge();

    // REMOVE_UNUSED_REGIONS : Option that remove section in the template that use data sources that does not exist
    // REMOVE_EMPTY_PARAGRAPHS : Option that remove empty lines in the generated document when data sources exist but is empty.
    // The deleted lines will only be those containing empty "merge fields", ":" and spaces.
    // If an alphanumeric character is present the paragraph will be generated.
    mailMerge.setCleanupOptions(MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS | MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS | MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS);

    // Add callback to enable html text in merge field
    // (http://www.aspose.com/community/forums/thread/380671/html-text-with-merge-field.aspx)
    mailMerge.setMergeDuplicateRegions(true);
    mailMerge.setRestartListsAtEachSection(true);

    // Authorize mustache syntax in word template
    mailMerge.setUseNonMergeFields(true);

    // Authorize to use the same region with same name multiple times
    mailMerge.setMergeDuplicateRegions(true);
    mailMerge.executeWithRegions(dataSourceRoot);

    // Fusion object is a global object that contains all data sources.
    // Send "fusion" object authorize to use data sources directly
    // with simple syntaxe like this "MERGERFIELD dossierFinancement.reference" or much simpler "{{ dossierFinancement.reference }}"
    mailMerge.execute(dataSourceRoot.getDataSource("fusion"));

    // Close template stream
    templateInput.close();

    doc.updateFields();

    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    doc.save(outputStream, new PdfSaveOptions());

    return outputStream;
  }
}
