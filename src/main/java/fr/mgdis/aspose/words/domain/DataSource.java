package fr.mgdis.aspose.words.domain;

import fr.mgdis.aspose.words.domain.Format.MIMEType;
import jakarta.validation.Valid;
import jakarta.validation.constraints.NotEmpty;
import jakarta.validation.constraints.NotNull;

/**
 * Record that represent a data source.
 *
 * @param name     Name of the data source. Use like identifier in template document
 * @param mimeType Content type of the date source
 * @param content  Content of the data-source
 */
@Valid
public record DataSource(@NotEmpty String name, @NotNull MIMEType mimeType, @NotEmpty String content) {
}
