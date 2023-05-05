package com.codinger.onlinePreview;

/**
 * @Author rongjia.wang
 * @Date 2023/4/23 21:49
 */

import com.spire.doc.Document;
import com.spire.presentation.Presentation;
import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Controller
@RequestMapping("/preview")
public class PreviewController {

    @GetMapping(value = "/xls")
    public ResponseEntity preview() throws IOException {
        Resource resource = new ClassPathResource("static/" + "1.xlsx");
        InputStream inputStream = resource.getInputStream();

        //Create a Workbook instance
        Workbook workbook = new Workbook();
        //Load an Excel file
        workbook.loadFromStream(inputStream);

        //Save the file to HTML
        workbook.saveToFile("ToHtml.html", FileFormat.HTML);

        return ResponseEntity.ok("");
    }

    @GetMapping(value = "/doc")
    public ResponseEntity preview1() throws IOException {
        Resource resource = new ClassPathResource("static/" + "1.docx");
        InputStream inputStream = resource.getInputStream();

        //Create a Document instance
        Document document = new Document();
        //Load a Word document
        document.loadFromFile(resource.getFile().getAbsolutePath());

        //Save the document as HTML
        document.saveToFile("output/toHtml.html");

        return ResponseEntity.ok("");
    }

    @GetMapping(value = "/ppt")
    public ResponseEntity preview2() throws Exception {
        Resource resource = new ClassPathResource("static/" + "1.pptx");
        InputStream inputStream = resource.getInputStream();

        //Create a Presentation object
        Presentation presentation = new Presentation();

        //Load the sample document
        presentation.loadFromFile(resource.getFile().getAbsolutePath());

        //Save the document to HTML format
        presentation.saveToFile("output/ppt.html", com.spire.presentation.FileFormat.HTML);

        return ResponseEntity.ok("");
    }

    @GetMapping("/download")
    public void download(HttpServletResponse response) throws IOException {
        // List of file names to compress
        Resource resource1 = new ClassPathResource("static/" + "1.pptx");
        Resource resource2 = new ClassPathResource("static/" + "1.xlsx");
        String[] filePaths = {resource1.getFile().getAbsolutePath(), resource2.getFile().getAbsolutePath()};

        // Set the response headers
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=\"compressed.zip\"");

        // Create a ZipOutputStream to hold the compressed data
        ZipOutputStream zos = new ZipOutputStream(response.getOutputStream());

        // Loop through the file paths and add each file to the compressed output
        for (String filePath : filePaths) {
            // Create a Path object for the current file
            Path path = Paths.get(filePath);

            // Create a ZipEntry for the current file and add it to the ZipOutputStream
            ZipEntry zipEntry = new ZipEntry(path.getFileName().toString());
            zos.putNextEntry(zipEntry);

            // Write the contents of the file to the ZipOutputStream
            byte[] buffer = new byte[1024];
            int len;
            try (InputStream is = Files.newInputStream(path)) {
                while ((len = is.read(buffer)) > 0) {
                    zos.write(buffer, 0, len);
                }
            }

            // Close the current ZipEntry
            zos.closeEntry();
        }

        // Close the ZipOutputStream
        zos.close();
    }

}