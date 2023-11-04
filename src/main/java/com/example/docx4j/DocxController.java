package com.example.docx4j;

import org.docx4j.jaxb.Context;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;


@RestController
public class DocxController {

    private final DocxService docxService;

    @Value("${app.filename.exist}")
    private String existFilename;

    @Value("${app.filename.export}")
    private String exportFilename;

    @Value("${app.img}")
    private String img;

    @Autowired()
    public DocxController(DocxService docxService) {
        this.docxService = docxService;
    }


    @GetMapping("create-docx")
    public ResponseEntity<?> createDocx() throws Exception {
        WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();
        File exportFile = new File(exportFilename);


        //Text Elements and Styling
        mainDocumentPart.addStyledParagraphOfText("Title", "Hello World!");
        mainDocumentPart.addParagraphOfText("Welcome To Baeldung");
        ObjectFactory factory = Context.getWmlObjectFactory();
        P p = factory.createP();
        R r = factory.createR();
        Text t = factory.createText();
        t.setValue("Welcome To Baeldung");
        r.getContent().add(t);
        p.getContent().add(r);
        RPr rpr = factory.createRPr();
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        rpr.setB(b);
        rpr.setI(b);
        rpr.setCaps(b);
        Color green = factory.createColor();
        green.setVal("green");
        rpr.setColor(green);
        r.setRPr(rpr);
        mainDocumentPart.getContent().add(p);


        //Working With Images
        docxService.workWithImage(wordPackage, mainDocumentPart, img);

        //Create table
        docxService.createTable(wordPackage);

        // SAVE DOCX
        wordPackage.save(exportFile);

        return new ResponseEntity<>(HttpStatus.OK);
    }

    @GetMapping("traversal-replace")
    public ResponseEntity<?> traversalReplaceTxt() throws Docx4JException {
        WordprocessingMLPackage existingDoc  = WordprocessingMLPackage.load(new File(existFilename));
        existingDoc  = docxService.replaceTextWithTraversalUtil(existingDoc );
        WordprocessingMLPackage newDoc  = WordprocessingMLPackage.createPackage();

        // Ottieni il corpo (Body) dei documenti
        Body existingBody = existingDoc.getMainDocumentPart().getJaxbElement().getBody();
        Body newBody = newDoc.getMainDocumentPart().getJaxbElement().getBody();

        // Copia il contenuto dal documento esistente al nuovo documento
        newBody.getContent().addAll(existingBody.getContent());

        File exportFile = new File(exportFilename);

        newDoc.save(exportFile);

        return new ResponseEntity<>(HttpStatus.OK);
    }

    @GetMapping("/read-docx")
    public ResponseEntity<String> readDocx() throws JAXBException, Docx4JException {
        File doc = new File(existFilename);
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                .load(doc);
        MainDocumentPart mainDocumentPart = wordMLPackage
                .getMainDocumentPart();
        String textNodesXPath = "//w:t";
        List<Object> textNodes= mainDocumentPart
                .getJAXBNodesViaXPath(textNodesXPath, true);
        for (Object obj : textNodes) {
            Text text = (Text) ((JAXBElement) obj).getValue();
            String textValue = text.getValue();
            System.out.println(textValue);
        }
        return ResponseEntity.ok("ok");
    }

    @GetMapping("/download")
    public ResponseEntity<Resource> downloadFile() {
        File file = new File(exportFilename);

        if (file.exists()) {
            Resource resource = new FileSystemResource(file);

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + "welcome.docx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .body(resource);
        } else {
            return new ResponseEntity<>(HttpStatus.NOT_FOUND);
        }

    }


}
