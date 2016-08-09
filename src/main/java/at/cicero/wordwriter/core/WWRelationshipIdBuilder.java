package at.cicero.wordwriter.core;

import java.io.IOException;

/**
 * Created by scd on 13.07.2016.
 */
public class WWRelationshipIdBuilder {
    public static Integer relIdCounter = WWDefaults.rId;

    public WWRelationshipIdBuilder() throws IOException {
    }

    public static String getRelId() {
        String relId = new String(WWDefaults.relationshipIdString);
        String result = relId + Integer.toString(relIdCounter);
        relIdCounter++;
        return result;
    }

    public static Integer getRelIdCounter() {
        Integer relId = relIdCounter;
        relIdCounter++;
        return relId;
    }
}
