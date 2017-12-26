package com.capgemini.license;

import java.util.ArrayList;
import java.util.List;

/**
 * Class that contains the information from each Main Application
 */
public class MainApplication {

    private final String name;

    private final String version;

    private final String description;

    private List<Component> components;

    /**
     * @param name
     * @param version
     * @param description
     */
    public MainApplication(String name, String version, String description) {
        this.name = name;
        this.version = version;
        this.description = description;
        components = new ArrayList<Component>();
    }

    public String getName() {
        return name;
    }

    public String getVersion() {
        return version;
    }

    public String getDescription() {
        return description;
    }

}
