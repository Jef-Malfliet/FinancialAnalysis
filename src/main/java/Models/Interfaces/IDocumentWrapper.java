package Models.Interfaces;

import Models.Business;
import Models.Enums.PropertyName;

import java.util.Map;

public interface IDocumentWrapper {

    int getYear();

    Business getBusiness();

    Map<PropertyName, String> getPropertiesMap();
}
