package main.test.powerPoint;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class PowerPointWithEasyRuns {
    private final String filePath;
    private final List<String> words;

    private final SeparateRunsMaker separateRunsMaker;
    private final WordsWithDifferentCasesProvider wordsWithDifferentCasesProvider;
    private final RunListProvider runListProvider;
    private final PowerPointSaver saver;
    private final RunListMapGenerator runListMapGenerator;

    private final XMLSlideShow slideShow;
    private final List<XSLFTextRun> runList;
    private final List<XSLFTextRun> soughtRuns;
    private final Map<String, Integer> soughtRunsMap;

    public static void main(String[] args) {
        PowerPointWithEasyRuns powerPointWithEasyRuns = new PowerPointWithEasyRuns("1.pptx", Arrays.asList("word"));
        System.out.println(powerPointWithEasyRuns.getSoughtRunsMap());
    }

    public PowerPointWithEasyRuns(String filePath, List<String> words) {
        separateRunsMaker = new SeparateRunsMaker();
        wordsWithDifferentCasesProvider = new WordsWithDifferentCasesProvider();
        runListProvider = new RunListProvider();
        saver = new PowerPointSaverImpl();
        runListMapGenerator = new RunListMapGenerator();

        this.filePath = filePath;
        this.words = words;

        XMLSlideShow tempSlideShow = null;
        try(FileInputStream is = new FileInputStream(filePath)) {
            tempSlideShow = new XMLSlideShow(is);
        } catch (IOException e) {
            e.printStackTrace();
        }
        slideShow = tempSlideShow;

        List<String> wordsDC = wordsWithDifferentCasesProvider.getWordsWithDifferentCases(words);

        separateRunsMaker.makeSeparateRuns(slideShow, wordsDC);

        runList = runListProvider.getRunList(slideShow);

        soughtRuns = runList.stream().filter(run -> wordsDC.contains(run.getRawText()))
                .collect(Collectors.toList());

        soughtRuns.forEach(run->run.setFontColor(Color.YELLOW));

        soughtRunsMap = runListMapGenerator.generateMap(soughtRuns);

        saver.save(slideShow, filePath, new String[]{String.valueOf(soughtRuns.size())});

    }

    public String getFilePath() {
        return filePath;
    }

    public List<String> getWords() {
        return words;
    }

    public XMLSlideShow getSlideShow() {
        return slideShow;
    }

    public List<XSLFTextRun> getRunList() {
        return runList;
    }

    public List<XSLFTextRun> getSoughtRuns() {
        return soughtRuns;
    }

    public Map<String, Integer> getSoughtRunsMap() {
        return soughtRunsMap;
    }
}
