package org.jozif;

import java.util.Objects;
import java.util.Set;

public class TaskUnit {
    private Integer excelRowNumber;// excel行号
    private String value;//excel单元格内容
    private String refinedValues;//excel替换掉字符，空格的单元格内容
    private Set<String> translatedValuesSet;//通过规则转换之后的新单词
    private Set<String> resultValuesSet;

    public TaskUnit(Integer excelRowNumber, String value) {
        this.excelRowNumber = excelRowNumber;
        this.value = value;
    }

    @Override
    public String toString() {
        return "TaskUnit{" +
                "excelRowNumber=" + excelRowNumber +
                ", value='" + value + '\'' +
                ", refinedValues='" + refinedValues + '\'' +
                ", translatedValuesSet=" + translatedValuesSet +
                ", resultValuesSet=" + resultValuesSet +
                '}';
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof TaskUnit)) return false;
        TaskUnit taskUnit = (TaskUnit) o;
        return getValue().equals(taskUnit.getValue());
    }

    @Override
    public int hashCode() {
        return Objects.hash(value);
    }

    public String getRefinedValues() {
        return refinedValues;
    }

    public void setRefinedValues(String refinedValues) {
        this.refinedValues = refinedValues;
    }

    public Integer getExcelRowNumber() {
        return excelRowNumber;
    }

    public void setExcelRowNumber(Integer excelRowNumber) {
        this.excelRowNumber = excelRowNumber;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public Set<String> getTranslatedValuesSet() {
        return translatedValuesSet;
    }

    public void setTranslatedValuesSet(Set<String> translatedValuesSet) {
        this.translatedValuesSet = translatedValuesSet;
    }

    public Set<String> getResultValuesSet() {
        return resultValuesSet;
    }

    public void setResultValuesSet(Set<String> resultValuesSet) {
        this.resultValuesSet = resultValuesSet;
    }
}
