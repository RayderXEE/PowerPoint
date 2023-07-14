package main.test.powerPoint;

import main.test.powerPoint.PowerPointSaver;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

import java.io.FileOutputStream;

public class PowerPointSaverImpl implements PowerPointSaver {

    @Override
    public void save(XMLSlideShow slideShow, String filePath, String[] args) {
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(filePath.replace(".pptx",
                    " (changed " + args[0] + " ).pptx"));
            slideShow.write(fileOutputStream);
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
