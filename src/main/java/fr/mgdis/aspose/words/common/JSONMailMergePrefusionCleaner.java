package fr.mgdis.aspose.words.common;

import io.quarkus.logging.Log;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

public class JSONMailMergePrefusionCleaner {
  private JSONObject _jsonObject;
  private JSONArray _jsonArray;

  private static final String FUSION = "fusion";

  public JSONMailMergePrefusionCleaner(String json) throws JSONException {
    if (json == null || json.isEmpty()) {
      json = "{}";
    }
    if (json.startsWith("[")) {
      this._jsonArray = new JSONArray(json);
    } else {
      this._jsonObject = new JSONObject(json);
    }
  }

  /**
   * Encapsulate json into a root element named "fusion"
   * It also applies specific process to the json
   *
   * @return String JSON
   */
  public String getCleanJson() {
    JSONObject root = new JSONObject();
    try {
      if (_jsonObject == null) {
        root.put(FUSION, _jsonArray);
      } else {
        root.put(FUSION, _jsonObject);
      }
    } catch (JSONException e) {
      // Just a log
      Log.error("getCleanJson has failed, original Json will be used instead", e);
    }

    return root.toString();
  }

}
