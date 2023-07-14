package main.test.powerPoint;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class PrettySlideShow {
    private final SeparateRunsMaker separateRunsMaker;
    private final WordsWithDifferentCasesProvider wordsWithDifferentCasesProvider;
    private final RunListProvider runListProvider;
    private final RunListMapGenerator runListMapGenerator;

    private final String filePath;
    private final List<String> words;
    private final XMLSlideShow slideShow;
    private final List<XSLFTextRun> runList;
    private final List<XSLFTextRun> soughtRuns;

    public PrettySlideShow(String filePath, List<String> words) {
        separateRunsMaker = new SeparateRunsMaker();
        wordsWithDifferentCasesProvider = new WordsWithDifferentCasesProvider();
        runListProvider = new RunListProvider();
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
}
