package fr.mgdis.aspose.words;

import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.ImageFieldMergingArgs;
import org.apache.commons.lang3.StringUtils;

/**
 * Override field merging to enable html text in merge field
 */
public class HandleMergeFields implements IFieldMergingCallback {
  /**
   * Method that allows you to add behavior or modify certain specific aspects of the document merge operation.
   *
   * @param fieldMergingArgs Field configuration of the merge field
   * @throws Exception Exception during field merging
   */
  @Override
  public void fieldMerging(FieldMergingArgs fieldMergingArgs) throws Exception {
    // Check if merge field contains a "\exists" flag
    if (fieldMergingArgs.getField().getFieldCode().contains("\\exists")) {
      processExistsFlag(fieldMergingArgs);
    }
    // Check if merge field contains a "\html" flag
    if (fieldMergingArgs.getField().getFieldCode().contains("\\html")) {
      processHTMLFlag(fieldMergingArgs);
    }
    // Check other "null" value
    if (fieldMergingArgs.getFieldValue() != null && fieldMergingArgs.getFieldValue().equals("null")) {
      fieldMergingArgs.setText("");
    }
  }

  private void processExistsFlag(FieldMergingArgs e) throws Exception {
    // Create document builder
    DocumentBuilder builder = new DocumentBuilder(e.getDocument());

    // Move cursor to field
    builder.moveToField(e.getField(), true);
    if (e.getFieldValue() != null && !e.getFieldValue().equals("null")) {
      String value = !e.getFieldValue().toString().isEmpty() ? "true" : "false";

      // Setting the flag in the document
      e.setText(value);
    }
  }

  private void processHTMLFlag(FieldMergingArgs e) throws Exception {
    // Then we add the style to the container to force the font inherit
    String htmlValue = (e.getFieldValue() != null && !e.getFieldValue().equals("null"))
      ? e.getFieldValue().toString()
      : StringUtils.EMPTY;
    // And after that we add the HTML Value in the document

    // Create document builder
    DocumentBuilder builder = new DocumentBuilder(e.getDocument());

    // Move cursor to field
    builder.moveToField(e.getField(), true);

    // Insert HTML
    builder.insertHtml(htmlValue, true);

    // Remove the string representation of HTML
    e.setText(StringUtils.EMPTY);
  }

  @Override
  public void imageFieldMerging(ImageFieldMergingArgs arg0) {
    // Auto-generated method stub
  }
}
