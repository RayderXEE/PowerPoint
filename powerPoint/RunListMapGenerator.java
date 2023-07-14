package main.test.powerPoint;

import org.apache.poi.xslf.usermodel.XSLFTextRun;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class RunListMapGenerator {

    public Map<String, Integer> generateMap(List<XSLFTextRun> runList) {
        Map<String, Integer> map = new HashMap<>();

        for (XSLFTextRun run : runList) {
            String word = run.getRawText();

            if (map.containsKey(word)) {
                map.put(word, map.get(word)+1);
            } else {
                map.put(word, 1);
            }
        }

        return map;

    }
}
