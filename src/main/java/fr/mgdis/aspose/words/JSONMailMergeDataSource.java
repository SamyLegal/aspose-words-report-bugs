package fr.mgdis.aspose.words;

import com.aspose.words.IMailMergeDataSource;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Implements a datasource for aspose word merging based on some json parsed with jackson tree parsing
 *
 * @see <a href="https://www.aspose.com/docs/display/wordsjava/IMailMergeDataSource">IMailMergeDataSource</a>
 * @see <a href="https://www.studytrails.com/java/json/java-jackson-json-tree-parsing.jsp">java-jackson-json-tree-parsing</a>
 */
public class JSONMailMergeDataSource implements IMailMergeDataSource {

  private static final ObjectMapper mapper = new ObjectMapper();
  private static final JsonNode emptyNode = mapper.createObjectNode();

  private String tableName;

  private Iterator<JsonNode> jsonNodeIterator;

  private JsonNode currentJsonNode;

  private JsonNode singleJsonNode;

  public JSONMailMergeDataSource(String tableName, String json) throws IOException {
    prepare(tableName, mapper.readTree(json));
  }

  public JSONMailMergeDataSource(String tableName, JsonNode jsonNode) {
    prepare(tableName, jsonNode);
  }

  public JSONMailMergeDataSource(String tableName) {
    prepare(tableName, null);
  }

  private void prepare(String tableName, JsonNode jsonNode) {
    this.tableName = tableName;
    // 3 possibilities :
    // - if the node is null or an empty array then it is initialized with an empty object
    // - it as if it was in an array of 1
    // - if it is an array we must be ready to iterate on it using 'moveNext'
    if (jsonNode == null || (jsonNode.isArray() && jsonNode.isEmpty())) {
      singleJsonNode = emptyNode;
    } else if (jsonNode.isArray()) {
      jsonNodeIterator = jsonNode.iterator();
    } else {
      singleJsonNode = jsonNode;
    }
  }

  @Override
  public String getTableName() {
    return this.tableName;
  }

  @Override
  public boolean getValue(String fieldName, Object[] fieldValue) {
    // The current JsonNode not exists
    if (currentJsonNode == null) return false;

    // When iterating over an array of simple values, return the value whatever the given field name
    if (!currentJsonNode.isObject() && !currentJsonNode.isArray()) {
      fieldValue[0] = currentJsonNode.asText();
      return true;
    }

    if (!currentJsonNode.has(fieldName)) {
      // If return false, the value will be «TableStart:table0»
      return true;
    }

    JsonNode child = currentJsonNode.get(fieldName);

    // a value can be a single element or an array
    if (child.isArray()) {
      List<String> items = new ArrayList<>();
      for (JsonNode jsonNode : child) {
        items.add(jsonNode.asText());
      }
      fieldValue[0] = items;
    } else {
      fieldValue[0] = child.asText();
    }

    return true;
  }

  @Override
  public boolean moveNext() {
    if (this.singleJsonNode != null && this.currentJsonNode != null) {
      // case where there was a single node given and it was already returned
      return false;
    } else if (this.singleJsonNode != null) {
      // case where there was a single node given and it has not yet been seen
      this.currentJsonNode = this.singleJsonNode;
      return true;
    } else if (this.jsonNodeIterator.hasNext()) {
      // case where we can iterate to the next child of the node given to the constructor
      this.currentJsonNode = this.jsonNodeIterator.next();
      return true;
    } else {
      return false;
    }
  }

  @Override
  public IMailMergeDataSource getChildDataSource(String tableName) {
    // Not sure if we should deal with recursivity (used for regions in aspose word merging)
    // in another case than a single root object
    if (currentJsonNode != null && currentJsonNode.has(tableName)) {
      return new JSONMailMergeDataSource(tableName, currentJsonNode.get(tableName));
    }
    return new JSONMailMergeDataSource(tableName);
  }
}
