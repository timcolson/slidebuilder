package com.feedblink;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.*;

import java.io.*;
import java.util.List;

public class Reader {
    public static void main(String[] args) {

        var filename = "source1-out";
        var r = new Reader();
        System.out.println("Read in a PPTX file: " + filename + ".pptx");
        try {
            r.read(filename);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void read(String file) throws IOException {
        System.out.println("Reading: " + file);

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("./" + file + ".pptx"));

        List<XSLFSlide> slides = ppt.getSlides();

        // Look at all the titles!
        for (XSLFSlide slide : slides) {
            System.out.println("slide = "
                    + slide.getTitle()
                    + "; name="
                    + slide.getSlideName()
            );
        }

        // Extract all image names
        // Note: in practice, also listed .wav and .mp3 files
        for (XSLFPictureData data : ppt.getPictureData()) {
            // doing nothing with the bytes of the file at present
            byte[] bytes = data.getData();
            String fileName = data.getFileName();
            // @Fixme - not the right object type here?
            PictureData.PictureType pictureFormat = data.getType();
            System.out.println("picture name: " + fileName);
            System.out.println("picture format: " + pictureFormat);
        }
        
        // Change Slide Order
        XSLFSlide slide = slides.get(2);
        ppt.setSlideOrder(slide, 1);

        // Add a new slide!
        // Get the slide master object
        XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);

        for( XSLFSlideLayout layout :  slideMaster.getSlideLayouts()) {
            System.out.println("layout = " + layout);
        }
        // Trouble getting by name, so just getting the first slide layout
        XSLFSlideLayout titleLayout = slideMaster.getSlideLayouts()[0];

        //creating a slide with title layout
        XSLFSlide slide4 = ppt.createSlide(titleLayout);

        //selecting the place holder in it
        XSLFTextShape title1 = slide4.getPlaceholder(0);
        //setting the title init
        title1.setText("Kilroy was here - title!");


        // Create an output object
        File targetFile = new File("./target/" + file + "-out.pptx");
        FileOutputStream out = new FileOutputStream(targetFile);
        // save modified PPT file
        ppt.write(out);
        System.out.println("Wrote to ./target/" + file + "-out.pptx !");
        out.close();


    }

}
