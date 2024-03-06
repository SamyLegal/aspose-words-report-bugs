package fr.mgdis.aspose.words;

import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.IMailMergeDataSourceRoot;
import fr.mgdis.aspose.words.domain.Format;
import fr.mgdis.aspose.words.domain.InputData;
import java.util.Collection;
import java.util.Optional;
import org.json.JSONException;

public class DataSourceRoot implements IMailMergeDataSourceRoot {
  private static final String INPUT_MIME_TYPE = "Input mime type";

  private final Collection<InputData> transformedDatasets;

  public DataSourceRoot(Collection<InputData> transformedDatasets) throws JSONException {
    this.transformedDatasets = transformedDatasets;
  }

  /**
   * Function call by Aspose during merge operation when it found a reference to a table name in a template.
   * Previously, data sources were stored in a map.
   * If you want to use a data source several times in a merge, you must return a new instance
   * of the "IMailMergeDataSource" type for it to work.
   *
   * @param tableName Name of the table. Define in the template with different syntax
   *                  Ex : {{#foreach dossierFinancement}} {{libelle}} {{/foreach dossierFinancement}}
   * @return Data Source
   */
  @Override
  public IMailMergeDataSource getDataSource(String tableName) throws Exception {
    Optional<InputData> inputData = transformedDatasets
      .stream()
      .filter(inputDataItem -> inputDataItem.getTableName().equals(tableName))
      .findAny();

    // If data source does not exist we return null with this value word not throw error and will not display this section
    if (inputData.isEmpty()) {
      return null;
    }

    var mimeType = inputData.get().getMimeType();
    if (mimeType == Format.MIMEType.JSON || mimeType == Format.MIMEType.HALJSON) {
      return new JSONMailMergeDataSource(tableName, inputData.get().getContent());
    }

    throw new UnsupportedOperationException(
      INPUT_MIME_TYPE + " " + mimeType + " not supported by Word document fusion"
    );
  }
}
