package fr.mgdis.aspose.words.config;

import io.quarkus.runtime.Startup;
import jakarta.enterprise.context.ApplicationScoped;

@Startup
@ApplicationScoped
public class AsposeWordsReportBugsConfig {

  public AsposeWordsReportBugsConfig() throws Exception {
    // Add license for Words
    com.aspose.words.License licenseWords = new com.aspose.words.License();
    licenseWords.setLicense(this.getClass().getResourceAsStream("/Aspose.Total.Java.lic"));
  }
}