package Interfaces;

import Domein.Business;
import Enums.PropertyName;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleStringProperty;

import java.util.Map;

public interface IDocumentWrapper {

    String getName();

    int getYear();

    Business getBusiness();

    SimpleStringProperty getNameProperty();

    BooleanProperty getSelectedProperty();

    Map<PropertyName, String> getPropertiesMap();

}
