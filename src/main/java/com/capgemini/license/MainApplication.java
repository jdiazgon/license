package com.capgemini.license;

/**
 * Class that contains the information from each Main Application
 */
public class MainApplication {

    private final String name;

    private final String version;

    private final String description;

    /**
     * @param name
     * @param version
     * @param description
     */
    public MainApplication(String name, String version, String description) {
        this.name = name;
        this.version = version;
        this.description = description;
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
