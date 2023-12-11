package fr.mgdis.aspose.words;

import jakarta.inject.Inject;
import jakarta.ws.rs.GET;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.core.Response;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

@Path(value = "/aspose-words")
public class AsposeWordsApi {
  @Inject
  AsposeWordsBuilder asposeWordsBuilder;

  @GET
  @Path(value = "/document")
  public Response mergeDocument() throws FileNotFoundException {
    try (InputStream template = new FileInputStream("src/test/resources/templates/template_01.docx");) {
      String content = new String(Files.readAllBytes(Paths.get("src/test/resources/data/data_01.json")));

      // IMailMergeDataSourceRoot that contains our data for the merge operation
      var dataSourceRoot = new JSONMailMergeDataSourceRoot(content);
      var document = asposeWordsBuilder.buildDocument(template, dataSourceRoot);

      return Response.ok(document.toByteArray())
          .header("Content-Type", "application/pdf")
          .header(
              "Content-Disposition",
              "attachment; filename=document_01.pdf"
          )
          .build();
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }
}
