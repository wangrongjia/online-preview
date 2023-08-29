package com.codinger.onlinePreview;

/**
 * @Author rongjia.wang
 * @Date 2023/4/23 21:49
 */

import cn.hutool.core.img.ImgUtil;
import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;
import fr.opensagres.poi.xwpf.converter.xhtml.Base64EmbedImgManager;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.http.*;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.w3c.dom.Document;

import javax.servlet.http.HttpServletResponse;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;


@Controller
@RequestMapping("/preview")
public class PreviewController {

    @GetMapping(value = "/excel")
    public ResponseEntity preview() throws IOException {
        Resource resource = new ClassPathResource("static/" + "1.xlsx");
        InputStream inputStream = resource.getInputStream();
        File htmlFile = new File("static/excel.html");

        //Create a Workbook instance
        Workbook workbook = new Workbook();
        //Load an Excel file
        workbook.loadFromStream(inputStream);

        //Save the file to HTML
        workbook.saveToFile("static/excel.html", FileFormat.HTML);
        return ResponseEntity.ok("file converted, path is static/excel.html");

    }

    @GetMapping(value = "/docx")
    public ResponseEntity previewDocx() throws IOException {
        Resource resource = new ClassPathResource("static/" + "1.docx");
        InputStream inputStream = resource.getInputStream();

        File htmlFile = new File("static/docx.html");

        XWPFDocument document = new XWPFDocument(inputStream);
        try (OutputStream out = Files.newOutputStream(htmlFile.toPath())) {
            XHTMLOptions options = XHTMLOptions.create().indent(4).setImageManager(new Base64EmbedImgManager());
//            options.setCharset("UTF-8"); // 设置编码方式为 UTF-8
            XHTMLConverter.getInstance().convert(document, out, options);
        }

        return ResponseEntity.ok("file converted, path is " + htmlFile);
    }

    @GetMapping(value = "/doc")
    public ResponseEntity preview1() throws IOException {
        Resource resource = new ClassPathResource("static/" + "1.doc");
        InputStream inputStream = resource.getInputStream();

        File htmlFile = new File("static/doc.html");
        OutputStream outputStream = new FileOutputStream(htmlFile);

        HWPFDocument wordDocument = new HWPFDocument(inputStream);

        try {
            convertDoc2Html(wordDocument, outputStream);
        } catch (ParserConfigurationException e) {
            throw new RuntimeException(e);
        } catch (TransformerException e) {
            throw new RuntimeException(e);
        }
        // 读取生成的 HTML 文件内容
        return ResponseEntity.ok("file converted, path is " + htmlFile);
    }

    private static void convertDoc2Html(HWPFDocument wordDocument, OutputStream outStream) throws ParserConfigurationException, TransformerException {
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder()
                        .newDocument());

        // 附件内含有图片 将图片内容转化成base64放到html文件中
        handleWordFileImage(wordToHtmlConverter);

        // 解析word文档
        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();

        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(outStream);

        TransformerFactory factory = TransformerFactory.newInstance();
        Transformer serializer = factory.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");

        serializer.transform(domSource, streamResult);
    }

    // 附件内含有图片 将图片内容转化成base64放到html文件中
    private static void handleWordFileImage(WordToHtmlConverter wordToHtmlConverter) {
        wordToHtmlConverter.setPicturesManager((content, pictureType, suggestedName, widthInches, heightInches) -> {
            BufferedImage bufferedImage = null;
            try {
                bufferedImage = ImgUtil.toImage(content);
            } catch (Exception e) {
                return "";
            }
            String base64Img = ImgUtil.toBase64(bufferedImage, pictureType.getExtension());
            //  带图片的word，则将图片转为base64编码，保存在一个页面中
            return "data:;base64," + base64Img;
        });
    }

    @GetMapping(value = "/ppt")
    public ResponseEntity preview2() throws Exception {
        Resource resource = new ClassPathResource("static/" + "1.pptx");

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