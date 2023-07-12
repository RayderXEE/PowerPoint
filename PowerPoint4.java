package main.test;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class PowerPoint4 {

    private XMLSlideShow slideShow;

    private String slideShowFilePath;
    private List<XSLFTextRun> foundRunList = new ArrayList<>();
    private int foundWordCount =0;
    private Map<String, Integer> foundWordMap = new HashMap<>();

    public static void main(String[] args) {

        try {
            String filePath = "1.pptx";

            PowerPoint4 powerPoint = new PowerPoint4(filePath);
            powerPoint.makeSeparateRuns("word");

            for (XSLFTextRun run : powerPoint.foundRunList) {
                run.setFontColor(Color.YELLOW);
            }

            System.out.println(powerPoint);

            powerPoint.save();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public PowerPoint4(String slideShowFilePath) {
        try {
            this.slideShowFilePath = slideShowFilePath;
            slideShow = new XMLSlideShow(new FileInputStream(slideShowFilePath));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void makeSeparateRuns(String searchWord) {
        List<String> words2 = new ArrayList<>();
        words2.add(searchWord);
        if (!containsHanScript(searchWord)) {
            if (!StringUtils.isAllUpperCase(searchWord)) {
                words2.add(StringUtils.capitalize(searchWord));
                words2.add(StringUtils.upperCase(searchWord));
            }
        }

        for (String word : words2) {
            for (XSLFSlide slide : slideShow.getSlides()) {
                for (int i = 0; i < slide.getShapes().size(); i++) {
                    XSLFShape shape = slide.getShapes().get(i);
                    if (shape instanceof XSLFTextShape) {
                        handleTextShape(word, (XSLFTextShape) shape);
                    }
                    if (shape instanceof XSLFTable) {
                        handleTable(word, (XSLFTable) shape);
                    }
                }
            }
        }
    }

    private void handleTable(String word, XSLFTable table) {
        for (XSLFTableRow row : table.getRows()) {
            for (XSLFTableCell cell : row.getCells()) {
                if (cell.getText().contains(word)) {
                    for (XSLFTextParagraph paragraph : cell.getTextParagraphs()) {
                        handleParagraph(word, paragraph);
                    }
                }
            }
        }
    }

    private void handleTextShape(String word, XSLFTextShape textBox) {
        for (XSLFTextParagraph paragraph : textBox.getTextParagraphs()) {
            handleParagraph(word, paragraph);
        }
    }

    private void handleParagraph(String word, XSLFTextParagraph paragraph) {
        List<Integer> foundWordIndexList = findWord(paragraph.getText(), word);

        for (Integer index : foundWordIndexList) {
            int index1 = index;
            int index2 = index1 + word.length();

            splitParagraph(paragraph, index1);
            splitParagraph(paragraph, index2);
        }

        for (XSLFTextRun run : paragraph.getTextRuns()) {
            if (run.getRawText().equals(word)) {
                foundRunList.add(run);
                foundWordCount++;
                if (foundWordMap.containsKey(word)) {
                    foundWordMap.put(word, foundWordMap.get(word)+1);
                } else {
                    foundWordMap.put(word, 1);
                }
            }
        }
    }

    private List<Integer> findWord(String text, String word) {
        ArrayList<Integer> list = new ArrayList<>();
        int index=-1;
        while ((index = text.indexOf(word, index)) != -1) {
            list.add(index);
            index++;
        }
        return list;
    }

    public void save() throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(slideShowFilePath.replace(".pptx",
                " (changed " + foundWordCount + " ).pptx"));
        slideShow.write(fileOutputStream);
        fileOutputStream.close();
    }

    private void copyRunsInParagraph(XSLFTextParagraph paragraph, XSLFTextRun stopRun) {
        paragraph.addNewTextRun();
        List<XSLFTextRun> runs = paragraph.getTextRuns();
        for (int j = runs.size()-2; j > -1; j--) {
            XSLFTextRun run = runs.get(j);
            copy(run, runs.get(j+1));
            if (run == stopRun) break;
        }
    }

    private void splitParagraph(XSLFTextParagraph paragraph, int splitIndex) {
        int index = 0;
        List<XSLFTextRun> textRuns = paragraph.getTextRuns();
        for (int i = 0; i < textRuns.size(); i++) {
            XSLFTextRun run = textRuns.get(i);
            int startRun = index;
            int endRun = index + run.getRawText().length();

            if (splitIndex >= startRun && splitIndex < endRun) {
                copyRunsInParagraph(paragraph, run);
                run.setText(run.getRawText().substring(0, splitIndex-startRun));
                XSLFTextRun run2 = paragraph.getTextRuns().get(i + 1);
                run2.setText(run2.getRawText().substring(splitIndex-startRun));
                return;
            }
            index = endRun;
        }
    }

    private void copy(XSLFTextRun from, XSLFTextRun to) {
        to.setText(from.getRawText());
        to.setFontColor(from.getFontColor());
        to.setFontSize(from.getFontSize());
        to.setFontFamily(from.getFontFamily());
        to.setBold(from.isBold());
        to.setItalic(from.isItalic());
        to.setUnderlined(from.isUnderlined());
        to.setStrikethrough(from.isStrikethrough());
        to.setSubscript(from.isSubscript());
        to.setSuperscript(from.isSuperscript());
        to.setCharacterSpacing(from.getCharacterSpacing());
    }

    public boolean containsHanScript(String s) {
        return s.codePoints().anyMatch(
                codepoint ->
                        Character.UnicodeScript.of(codepoint) == Character.UnicodeScript.HAN);
    }

    public XMLSlideShow getSlideShow() {
        return this.slideShow;
    }

    public String getSlideShowFilePath() {
        return this.slideShowFilePath;
    }

    public List<XSLFTextRun> getFoundRunList() {
        return this.foundRunList;
    }

    public int getFoundWordCount() {
        return this.foundWordCount;
    }

    public Map<String, Integer> getFoundWordMap() {
        return this.foundWordMap;
    }

    public String toString() {
        return "slideShowFilePath=" + this.getSlideShowFilePath() + ", foundWordCount=" + this.getFoundWordCount() + ", foundWordMap=" + this.getFoundWordMap() + ")";
    }

    private List<String> getFoundRunTextList() {
        List<String> list = new ArrayList<>();
        for (XSLFTextRun run : getFoundRunList()) {
            list.add(run.getRawText());
        }
        return list;
    }
}
