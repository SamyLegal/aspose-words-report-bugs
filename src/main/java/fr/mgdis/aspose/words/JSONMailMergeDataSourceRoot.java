package fr.mgdis.aspose.words;

import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.IMailMergeDataSourceRoot;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;

/**
 * Implements a datasource root for aspose word merging based on some json parsed with jackson tree parsing
 *
 * @see <a href="https://www.aspose.com/docs/display/wordsjava/IMailMergeDataSource">IMailMergeDataSource</a>
 * @see <a href="https://www.studytrails.com/java/json/java-jackson-json-tree-parsing.jsp">java-jackson-json-tree-parsing</a>
 */
public class JSONMailMergeDataSourceRoot implements IMailMergeDataSourceRoot {

  private static final ObjectMapper mapper = new ObjectMapper();

  private final JsonNode jsonNode;

  public JSONMailMergeDataSourceRoot(String json) throws IOException {
    jsonNode = mapper.readTree(json);
  }

  public IMailMergeDataSource getDataSource(String tableName) {
    if (!jsonNode.has(tableName)) {
      return new JSONMailMergeDataSource(tableName);
    }
    JsonNode child = jsonNode.get(tableName);
    return new JSONMailMergeDataSource(tableName, child);
  }
}
