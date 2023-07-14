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

public class PowerPointWithEasyRuns extends PrettySlideShow {

    private final PowerPointSaver saver;
    private final RunListMapGenerator runListMapGenerator;

    private final Map<String, Integer> soughtRunsMap;

    public static void main(String[] args) {
        PowerPointWithEasyRuns powerPointWithEasyRuns = new PowerPointWithEasyRuns("1.pptx", Arrays.asList("word"));
        System.out.println(powerPointWithEasyRuns.getSoughtRunsMap());
    }

    public PowerPointWithEasyRuns(String filePath, List<String> words) {
        super(filePath, words);
        saver = new PowerPointSaverImpl();
        runListMapGenerator = new RunListMapGenerator();

        getSoughtRuns().forEach(run->run.setFontColor(Color.YELLOW));

        soughtRunsMap = runListMapGenerator.generateMap(getSoughtRuns());

        saver.save(getSlideShow(), filePath, new String[]{String.valueOf(getSoughtRuns().size())});

    }

    public Map<String, Integer> getSoughtRunsMap() {
        return soughtRunsMap;
    }
}
