package fr.mgdis.aspose.words;

import io.quarkus.test.junit.QuarkusTest;
import jakarta.inject.Inject;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.assertTrue;

@QuarkusTest
class AsposeWordsBugTest {

  @Inject
  AsposeWordsBuilder asposeWordsBuilder;

  @Test
  @DisplayName("should merge document with old word template and new version of aspose words without throw an exception")
  void testWordMergeOperationWithAnOldWordTemplate() throws Exception {
    try (InputStream template = new FileInputStream("src/test/resources/templates/template_01.docx");
         InputStream data = new FileInputStream("src/test/resources/data/data_01.json")) {

      String content = new String(Files.readAllBytes(Paths.get("src/test/resources/data/data_01.json")));

      // IMailMergeDataSourceRoot that contains our data for the merge operation
      var dataSourceRoot = new JSONMailMergeDataSourceRoot(content);
      var document = asposeWordsBuilder.buildDocument(template, dataSourceRoot);

      try (PDDocument pdfDocument = PDDocument.load(document.toByteArray())) {
        PDFTextStripper stripper = new PDFTextStripper();
        String text = stripper.getText(pdfDocument);

        // Check "Domiciliations bancaires"
        assertTrue(text.contains("Relevé d’identité bancaire : "));
        assertTrue(text.contains("word-template.docx - 07/12/2023 (19,34Ko)"));
        assertTrue(text.contains("Note : Mon template word. On peut avoir que un seul document"));

        // Check "Pièces"
        assertTrue(text.contains("Pièces fournies :"));
        assertTrue(text.contains("Pièce d'identité :"));
        assertTrue(text.contains("data-sources.pdf - 07/12/2023 (145,65Ko)"));
        assertTrue(text.contains("Note : Data Sources"));

        assertTrue(text.contains("Cahier des charges rédigé par le demandeur : Transmis par envoi postal"));

        assertTrue(text.contains("Propositions et devis des bureaux d'étude consultés :"));
        assertTrue(text.contains("word-template.docx - 07/12/2023 (19,34Ko)"));
        assertTrue(text.contains("Note : Mon template word"));

        assertTrue(text.contains("Pièce conditionnée au Plan de Financement TTC :"));
        assertTrue(text.contains("Toute pièce utile à la présentation de votre projet :"));
        assertTrue(text.contains("depot_TELE_TEST.docx - 07/12/2023 (50,94Ko)"));
      }
    }
  }
}