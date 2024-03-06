package fr.mgdis.aspose.words.domain;

/**
 * Utilities to deal with mime types, file formats, etc.
 */
public final class Format {

  private Format() {}

  /**
   * Supported input, output or templates mime types
   */
  public enum MIMEType {
    BIN("application/octet-stream"),
    XLS("application/vnd.ms-excel"),
    XLSX("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    XLSM("application/vnd.ms-excel.sheet.macroEnabled.12"),
    XLTX("application/vnd.openxmlformats-officedocument.spreadsheetml.template"),
    XLTM("application/vnd.ms-excel.template.macroEnabled.12"),
    XPS("application/vnd.ms-xpsdocument"),
    DOC("application/msword"),
    DOCX("application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
    PDF("application/pdf"),
    CSV("text/csv"),
    HTML("text/html"),
    ODS("application/vnd.oasis.opendocument.spreadsheet"),
    ODT("application/vnd.oasis.opendocument.text"),
    RTF("application/rtf"),
    JSON("application/json"),
    HALJSON("application/hal+json"),
    XML("application/xml"),
    EML("message/rfc822"),
    TXT("text/plain"),
    URI("text/uri-list"),
    JS("application/javascript");

    private final String value;

    MIMEType(String value) {
      this.value = value.toLowerCase();
    }

    public String getValue() {
      return this.value;
    }

    public String getFileExtension() {
      return this.toString().toLowerCase();
    }

    public static MIMEType fromValue(String value) {
      if (value.contains(";")) {
        value = value.split(";")[0];
      }
      if (value != null) {
        for (MIMEType type : MIMEType.values()) {
          if (value.equalsIgnoreCase(type.value)) {
            return type;
          }
        }
      }
      throw new IllegalArgumentException("No mime-type with value " + value + " supported");
    }

    /**
     * Return the mime type that corresponds to the given file extension
     * @param extension Extension file
     */
    public static MIMEType fromExtension(String extension) {
      if (extension != null) {
        for (MIMEType type : MIMEType.values()) {
          if (extension.equalsIgnoreCase(type.getFileExtension())) {
            return type;
          }
        }
      }
      throw new IllegalArgumentException("No mime-type found from extension " + extension);
    }
  }
}
