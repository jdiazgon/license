package com.capgemini.license;

/**
 *
 */
public class PlugIn {

    private final String name;

    private final String version;

    /**
     * @param name
     * @param version
     * @param description
     */
    public PlugIn(String name, String version) {
        this.name = name;
        this.version = version;
    }

    public String getName() {
        return name;
    }

    public String getVersion() {
        return version;
    }

}
