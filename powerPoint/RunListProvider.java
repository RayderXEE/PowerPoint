package main.test.powerPoint;

import org.apache.poi.xslf.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class RunListProvider {

    public List<XSLFTextRun> getRunList(XMLSlideShow slideShow) {
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

}
