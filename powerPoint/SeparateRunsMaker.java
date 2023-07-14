package main.test.powerPoint;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class SeparateRunsMaker {


    public static void main(String[] args) {

        try {
            String filePath = "1.pptx";
            List<String> wordList = new ArrayList<>();
            wordList.add("word");

            SeparateRunsMaker powerPoint = new SeparateRunsMaker();
            XMLSlideShow slideShow = new XMLSlideShow(new FileInputStream(filePath));

            powerPoint.makeSeparateRuns(slideShow, wordList);

            List<XSLFTextRun> runList = getRunList(slideShow);

            List<String> wordsWithDifferentCases = getWordsWithDifferentCases(wordList);

            List<XSLFTextRun> runList1 = new ArrayList<>();

            for (String word : wordsWithDifferentCases) {
                for (XSLFTextRun run : runList) {
                    if (run.getRawText().equals(word)) {
                        runList1.add(run);
                    }
                }
            }

            for (XSLFTextRun run : runList1) {
                run.setFontColor(Color.YELLOW);
                //System.out.println(run.getRawText());
            }

            PowerPointSaver saver = new PowerPointSaverImpl();
            saver.save(slideShow, filePath, new String[]{String.valueOf(runList1.size())});
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static List<XSLFTextRun> getRunList(XMLSlideShow slideShow) {
        List<XSLFTextRun> runList = new ArrayList<>();
        for (XSLFSlide slide : slideShow.getSlides()) {
            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape) shape;
                    for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                        runList.addAll(paragraph.getTextRuns());
                    }
                }
                if (shape instanceof XSLFTable) {
                    XSLFTable table = (XSLFTable) shape;
                    for (XSLFTableRow row : table.getRows()) {
                        for (XSLFTableCell cell : row) {
                            for (XSLFTextParagraph paragraph : cell.getTextParagraphs()) {
                                runList.addAll(paragraph.getTextRuns());
                            }

                        }

                    }

                }
            }

        }
        return runList;
    }

    public SeparateRunsMaker() {

    }

    public void makeSeparateRuns(XMLSlideShow slideShow, List<String> words) {
        for (String word : words) {
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

    private static List<String> getWordsWithDifferentCases(List<String> words) {
        List<String> words2 = new ArrayList<>();
        for (String word : words) {
            words2.add(word);
            if (!containsHanScript(word)) {
                if (!StringUtils.isAllUpperCase(word)) {
                    words2.add(StringUtils.capitalize(word));
                    words2.add(StringUtils.upperCase(word));
                }
            }
        }
        return words2;
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

    public static boolean containsHanScript(String s) {
        return s.codePoints().anyMatch(
                codepoint ->
                        Character.UnicodeScript.of(codepoint) == Character.UnicodeScript.HAN);
    }

}
