package com.capgemini.license;

/**
 * Embedded component data retrieved from the Excel file
 */
public class EmbeddedComponent {

    private final String name;

    private final String version;

    private final String description;

    /**
     * @param name
     * @param version
     * @param description
     */
    public EmbeddedComponent(String name, String version, String description) {
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
