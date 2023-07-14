package main.test.powerPoint;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public interface PowerPointSaver {
    void save(XMLSlideShow slideShow, String filePath, String[] args);
}
