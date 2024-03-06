package fr.mgdis.aspose.words.common;

import com.fasterxml.jackson.databind.ObjectMapper;
import fr.mgdis.aspose.words.domain.DataSource;
import fr.mgdis.aspose.words.domain.Format;
import fr.mgdis.aspose.words.domain.InputData;
import jakarta.enterprise.context.ApplicationScoped;
import java.io.IOException;
import java.io.InputStream;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.script.ScriptException;
import org.json.JSONException;
import org.json.JSONObject;

@ApplicationScoped
public class TransformationService {
  private static final String INPUT_MIME_TYPE = "Input mime type";

  ObjectMapper mapper;

  public TransformationService(ObjectMapper mapper) {
    this.mapper = mapper;
  }


  /**
   * Function that apply transformation script on global dataset
   *
   * @param transformationScript Transformation script
   * @param datasets             Datasets on each we apply this transformation
   * @return Return a list of "InputData" elements send directly as Aspose for the merge operation on the document.
   */
  public Collection<InputData> getTransformedDataSet(InputStream transformationScript, List<DataSource> datasets)
    throws JSONException, IOException, NoSuchMethodException, ScriptException {
    // Datasets with transformation
    Map<String, InputData> transformedDatasets = new HashMap<>();

    // Initialize a JSON entity with every request result (in a property "TABLE{n}").
    JSONObject jsonGlobalData = new JSONObject();

    for (DataSource dataSource : datasets) {
      Format.MIMEType mimeType = dataSource.mimeType();
      String data = dataSource.content();
      String tableName = dataSource.name();

      // We first initialize transformedDatasets with original values (case of no Transformation Script).
      transformedDatasets.put(tableName, new InputData(mimeType, tableName, data));

      if ((mimeType == Format.MIMEType.JSON) || (mimeType == Format.MIMEType.HALJSON)) {
        String cleanJson = new JSONMailMergePrefusionCleaner(data).getCleanJson();

        // Method getCleanJson insert a root node called "fusion" so we retrieve the content
        // of our data directly in the “fusion” node
        jsonGlobalData.put(tableName, new JSONObject(cleanJson).get("fusion"));
      } else if (mimeType == Format.MIMEType.CSV) {
        jsonGlobalData.put(tableName, data);
      } else {
        throw new UnsupportedOperationException(
          String.format("%s %s not supported by Word document fusion (datasource : \"%s\").", INPUT_MIME_TYPE,
            dataSource.mimeType().getValue(), tableName));
      }
    }

    // Use for compatibility between WordResource and DocumentResource
    // If we have a date source named "fusion" we use it directly
    String fusionNodeContent =
      jsonGlobalData.has("fusion") ? jsonGlobalData.get("fusion").toString() : jsonGlobalData.toString();

    // Compatibility purpose : in order to MS Word MergeField to work properly we have to
    // use a "root" node (a node that contains every TABLE{n}).
    // Such a node is usually labeled "fusion" in this application
    String globalTableName = "fusion";
    transformedDatasets.put(globalTableName, new InputData(Format.MIMEType.JSON, globalTableName, fusionNodeContent));
    jsonGlobalData.put(globalTableName, new JSONObject(fusionNodeContent));

    return transformedDatasets.values();
  }
}
