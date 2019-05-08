package org.jozif;

public class CustomizeRuleUnit {
    private String suffix;
    private String newSuffix;

    public CustomizeRuleUnit() {
    }

    public CustomizeRuleUnit(String suffix, String newSuffix) {
        this.suffix = suffix;
        this.newSuffix = newSuffix;
    }

    @Override
    public String toString() {
        return "CustomizeRuleUnit{" +
                "suffix='" + suffix + '\'' +
                ", newSuffix='" + newSuffix + '\'' +
                '}';
    }

    public String getSuffix() {
        return suffix;
    }

    public void setSuffix(String suffix) {
        this.suffix = suffix;
    }

    public String getNewSuffix() {
        return newSuffix;
    }

    public void setNewSuffix(String newSuffix) {
        this.newSuffix = newSuffix;
    }
}
