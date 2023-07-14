package main.test.powerPoint;

import org.apache.commons.lang.StringUtils;

import java.util.ArrayList;
import java.util.List;

public class WordsWithDifferentCasesProvider {
    public List<String> getWordsWithDifferentCases(List<String> wordList) {
        List<String> words2 = new ArrayList<>();
        for (String word : wordList) {
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

    public static boolean containsHanScript(String s) {
        return s.codePoints().anyMatch(
                codepoint ->
                        Character.UnicodeScript.of(codepoint) == Character.UnicodeScript.HAN);
    }

}
