package com.capgemini.license;

/**
 * Class containing the library data used in a main application or an embedded component. (name, version and
 * license)
 */
public class Component {

    private final String name;

    private final String version;

    private final String license;

    /**
     * @param name
     * @param version
     * @param description
     */
    public Component(String name, String version, String license) {
        this.name = name;
        this.version = version;
        this.license = license;
    }

    public String getName() {
        return name;
    }

    public String getVersion() {
        return version;
    }

    public String getLicense() {
        return license;
    }
}
