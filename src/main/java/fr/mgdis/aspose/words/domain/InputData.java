package fr.mgdis.aspose.words.domain;

import fr.mgdis.aspose.words.domain.Format.MIMEType;

public class InputData {
  private final MIMEType mimeType;

  private final String tableName;

  private String content;

  public InputData(MIMEType mimeType, String tableName, String content) {
    this.mimeType = mimeType;
    this.tableName = tableName;
    this.setContent(content);
  }

  public MIMEType getMimeType() {
    return mimeType;
  }

  public String getTableName() {
    return tableName;
  }

  public String getContent() {
    return content;
  }

  public void setContent(String content) {
    this.content = content;
  }
}
