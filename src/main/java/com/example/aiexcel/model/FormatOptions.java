package com.example.aiexcel.model;

/**
 * 格式选项模型
 * 用于定义Excel单元格的格式设置选项
 */
public class FormatOptions {
    
    // 字体属性
    private Boolean bold;
    private Boolean italic;
    private Integer fontSize;
    private String fontColor;
    
    // 背景和边框
    private String backgroundColor;
    private String borderColor;
    private String borderStyle; // solid, dashed, dotted, etc.
    
    // 对齐属性
    private String horizontalAlignment; // left, center, right
    private String verticalAlignment;   // top, middle, bottom
    
    // 数字格式
    private String numberFormat; // 0.00, 0%, etc.
    
    // 文本格式
    private Boolean wrapText;
    private Integer rotation;

    // Constructors
    public FormatOptions() {}
    
    public FormatOptions(Boolean bold, Boolean italic, Integer fontSize, String fontColor,
                         String backgroundColor, String borderColor, String borderStyle,
                         String horizontalAlignment, String verticalAlignment,
                         String numberFormat, Boolean wrapText, Integer rotation) {
        this.bold = bold;
        this.italic = italic;
        this.fontSize = fontSize;
        this.fontColor = fontColor;
        this.backgroundColor = backgroundColor;
        this.borderColor = borderColor;
        this.borderStyle = borderStyle;
        this.horizontalAlignment = horizontalAlignment;
        this.verticalAlignment = verticalAlignment;
        this.numberFormat = numberFormat;
        this.wrapText = wrapText;
        this.rotation = rotation;
    }

    // Getters and Setters
    public Boolean getBold() {
        return bold;
    }

    public void setBold(Boolean bold) {
        this.bold = bold;
    }

    public Boolean getItalic() {
        return italic;
    }

    public void setItalic(Boolean italic) {
        this.italic = italic;
    }

    public Integer getFontSize() {
        return fontSize;
    }

    public void setFontSize(Integer fontSize) {
        this.fontSize = fontSize;
    }

    public String getFontColor() {
        return fontColor;
    }

    public void setFontColor(String fontColor) {
        this.fontColor = fontColor;
    }

    public String getBackgroundColor() {
        return backgroundColor;
    }

    public void setBackgroundColor(String backgroundColor) {
        this.backgroundColor = backgroundColor;
    }

    public String getBorderColor() {
        return borderColor;
    }

    public void setBorderColor(String borderColor) {
        this.borderColor = borderColor;
    }

    public String getBorderStyle() {
        return borderStyle;
    }

    public void setBorderStyle(String borderStyle) {
        this.borderStyle = borderStyle;
    }

    public String getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public void setHorizontalAlignment(String horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
    }

    public String getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(String verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public String getNumberFormat() {
        return numberFormat;
    }

    public void setNumberFormat(String numberFormat) {
        this.numberFormat = numberFormat;
    }

    public Boolean getWrapText() {
        return wrapText;
    }

    public void setWrapText(Boolean wrapText) {
        this.wrapText = wrapText;
    }

    public Integer getRotation() {
        return rotation;
    }

    public void setRotation(Integer rotation) {
        this.rotation = rotation;
    }
}