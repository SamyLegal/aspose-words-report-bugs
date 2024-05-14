package fr.mgdis.aspose.words;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.adelean.inject.resources.junit.jupiter.GivenJsonResource;
import com.adelean.inject.resources.junit.jupiter.TestWithResources;
import com.fasterxml.jackson.databind.JsonNode;
import fr.mgdis.aspose.words.common.TransformationService;
import fr.mgdis.aspose.words.domain.DataSource;
import fr.mgdis.aspose.words.domain.Format;
import fr.mgdis.aspose.words.domain.InputData;
import io.quarkus.test.junit.QuarkusTest;
import jakarta.inject.Inject;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Collection;
import java.util.List;
import java.util.stream.StreamSupport;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

@QuarkusTest
@TestWithResources
class AsposeWordsBugTest {

  @Inject
  AsposeWordsBuilder asposeWordsBuilder;

  @Inject
  TransformationService transformationService;

  @Nested
  class Bug01 {
    @Test
    @DisplayName("should merge document with old word template and new version of aspose words without throw an exception")
    void testWordMergeOperationWithAnOldWordTemplate() throws Exception {
      try (InputStream template = new FileInputStream("src/test/resources/templates/template_01.docx")) {

        String content = new String(Files.readAllBytes(Paths.get("src/test/resources/data/data_01.json")));

        // IMailMergeDataSourceRoot that contains our data for the merge operation
        var dataSourceRoot = new JSONMailMergeDataSourceRoot(content);
        var document = asposeWordsBuilder.buildDocument(template, dataSourceRoot);

        try (PDDocument pdfDocument = PDDocument.load(document.toByteArray())) {
          PDFTextStripper stripper = new PDFTextStripper();
          String text = stripper.getText(pdfDocument);

          // Check "Domiciliations bancaires"
          assertTrue(text.contains("Relevé d’identité bancaire : "));
          // assertTrue(text.contains("word-template.docx - 07/12/2023 (19,34Ko)"));
          // assertTrue(text.contains("Note : Mon template word. On peut avoir que un seul document"));

          // Check "Pièces"
          assertTrue(text.contains("Pièces fournies :"));
          // assertTrue(text.contains("Pièce d'identité :"));
          // assertTrue(text.contains("data-sources.pdf - 07/12/2023 (145,65Ko)"));
          // assertTrue(text.contains("Note : Data Sources"));

          assertTrue(text.contains("Cahier des charges rédigé par le demandeur : Transmis par envoi postal"));

          assertTrue(text.contains("Propositions et devis des bureaux d'étude consultés :"));
          // assertTrue(text.contains("word-template.docx - 07/12/2023 (19,34Ko)"));
          // assertTrue(text.contains("Note : Mon template word"));

          assertTrue(text.contains("Pièce conditionnée au Plan de Financement TTC :"));
          // assertTrue(text.contains("Toute pièce utile à la présentation de votre projet :"));
          // assertTrue(text.contains("depot_TELE_TEST.docx - 07/12/2023 (50,94Ko)"));
        }
      }
    }
  }

  @Nested
  class Bug02 {
    @Test
    @DisplayName("should merge a document without range error when we call aspose words updateFields method")
    void testAsposeWordsRangeError(@GivenJsonResource("/data/data_02.json") JsonNode jsonNode) throws Exception {
      List<DataSource> dataSources = StreamSupport
        .stream(jsonNode.get("dataSources").spliterator(), false)
        .map(dataSourceNode -> new DataSource(dataSourceNode.get("name").asText(), Format.MIMEType.JSON,
          dataSourceNode.get("content").asText()))
        .toList();

      // Transformed datasets "datasets" on which the transformation functions have been applied
      Collection<InputData> transformedDatasets = transformationService.getTransformedDataSet(null, dataSources);

      try (InputStream template = new FileInputStream("src/test/resources/templates/template_02.docx")) {
        var document = asposeWordsBuilder.buildDocument(template, new DataSourceRoot(transformedDatasets));

        try (PDDocument pdfDocument = PDDocument.load(document.toByteArray())) {
          PDFTextStripper stripper = new PDFTextStripper();
          String text = stripper.getText(pdfDocument);

          // Check content
          assertTrue(text.contains(
            "Aides publiques perçues par l'organisme sur les 3 derniers exercices fiscaux (Dont l'année en cours) - "));
        }
      }
    }
  }

  @Nested
  class Bug03 {
    @Test
    @DisplayName("should merge a document with template of 'demande-financement recapitulatif' and data with these characters \" « »")
    void testMergeDocumentWithSpecialCharacters(@GivenJsonResource("/data/data_03.json") JsonNode jsonNode)
      throws Exception {
      List<DataSource> dataSources = StreamSupport
        .stream(jsonNode.get("dataSources").spliterator(), false)
        .map(dataSourceNode -> new DataSource(dataSourceNode.get("name").asText(), Format.MIMEType.JSON,
          dataSourceNode.get("content").asText()))
        .toList();

      // Transformed datasets "datasets" on which the transformation functions have been applied
      Collection<InputData> transformedDatasets = transformationService.getTransformedDataSet(null, dataSources);

      try (InputStream template = new FileInputStream("src/test/resources/templates/template_03.docx")) {
        var document = asposeWordsBuilder.buildDocument(template, new DataSourceRoot(transformedDatasets));
        assertNotNull(document);

        try (PDDocument pdfDocument = PDDocument.load(document.toByteArray())) {
          PDFTextStripper stripper = new PDFTextStripper();
          String text = stripper.getText(pdfDocument);

          // Check content
          assertTrue(text.contains("«  » \" Fiche d’identité CFA (indicateurs annuels)"));
        }
      }
    }
  }

  @Nested
  class Bug04 {
    @Test
    @DisplayName("should must generate a word document with the data as with an old version of Aspose")
    void testWordMergeOperationWithAnOldWordTemplate() throws Exception {
      try (InputStream template = new FileInputStream("src/test/resources/templates/template_04.docx")) {

        String content = new String(Files.readAllBytes(Paths.get("src/test/resources/data/data_04.json")));

        // IMailMergeDataSourceRoot that contains our data for the merge operation
        var dataSourceRoot = new JSONMailMergeDataSourceRoot(content);
        var document = asposeWordsBuilder.buildDocument(template, dataSourceRoot);

        try (PDDocument pdfDocument = PDDocument.load(document.toByteArray())) {
          PDFTextStripper stripper = new PDFTextStripper();
          String text = stripper.getText(pdfDocument);

          // Check "IBAN"
          assertTrue(text.contains("John DOE FR7730003000405359159988A02 AGRIFRPPXXX"));

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
}