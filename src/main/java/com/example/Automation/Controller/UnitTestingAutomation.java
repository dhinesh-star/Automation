package com.example.Automation.Controller;

import com.google.api.services.docs.v1.Docs;
import com.google.api.services.docs.v1.model.Document;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("unit-testing")
public class UnitTestingAutomation {
    @PostMapping("writeInDocs")
    public ResponseEntity writeInDocs(){
        try {
            String folderPath = "C:\\Users\\Ryzen 5\\OneDrive\\Documents\\Pictures\\Screenshots";
            List<File> fileList = getImagesFromFolder(folderPath);
            XWPFDocument document = new XWPFDocument();

            XWPFParagraph paragraph1 = document.createParagraph();
            paragraph1.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun run1 = paragraph1.createRun();
            run1.setBold(true);
            run1.setFontSize(20);
            run1.setText("Hello World");
            run1.addBreak();

            XWPFParagraph paragraph2 = document.createParagraph();
            paragraph2.setAlignment(ParagraphAlignment.LEFT);

            XWPFRun run2 = paragraph2.createRun();
            run2.setBold(true);
            run2.setText("Request ID: R256");
            run2.addBreak();
            run2.addBreak();

            XWPFParagraph paragraph3 = document.createParagraph();
            paragraph3.setAlignment(ParagraphAlignment.LEFT);

            for(File image:fileList){
                try(FileInputStream imageStream = new FileInputStream(image)){
                    XWPFRun run3 = paragraph3.createRun();

                    //Get the width and height of the image
                    BufferedImage bufferedImage = ImageIO.read(image);
                    int width = bufferedImage.getWidth();
                    int height = bufferedImage.getHeight();

                    int imageFormat = XWPFDocument.PICTURE_TYPE_PNG;
                    run3.addPicture(imageStream, imageFormat, image.getName(), Units.pixelToEMU(624), Units.pixelToEMU(350));
                    run3.addBreak();
                }catch (Exception e){
                    return new ResponseEntity<>(e.getMessage(),HttpStatus.BAD_REQUEST);
                }
            }

            try (FileOutputStream out = new FileOutputStream("unit_testing.doc")) {
                document.write(out);
            } catch (Exception e) {
                return new ResponseEntity<>(e.getMessage(), HttpStatus.BAD_REQUEST);
            }

            return new ResponseEntity<>("Successfully writen", HttpStatus.CREATED);
        }catch (Exception e) {
            return new ResponseEntity<>(e.getMessage(), HttpStatus.BAD_REQUEST);
        }
    }
    public static List<File> getImagesFromFolder(String folderPath){
        List<File> fileList = new ArrayList<>();
        File folder = new File(folderPath);

        if(folder.isDirectory()==true){
            File[] files = folder.listFiles();
            for(File file:files){
                if(isImageFile(file)==true){
                    fileList.add(file);
                }
            }
        }

        return fileList;
    }
    private static boolean isImageFile(File file) {
        // Check if the file has a valid image file extension
        String fileName = file.getName().toLowerCase();
        return fileName.endsWith(".jpg") || fileName.endsWith(".jpeg") ||
                fileName.endsWith(".png") || fileName.endsWith(".gif") ||
                fileName.endsWith(".bmp") || fileName.endsWith(".tiff");
    }
}
