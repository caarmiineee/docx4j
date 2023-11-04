package com.example.docx4j;

import lombok.Value;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.springframework.stereotype.Service;

import java.io.File;
import java.nio.file.Files;
import java.util.List;

@Service
public class DocxServiceImpl implements DocxService {
    @Override
    public WordprocessingMLPackage replaceTextWithTraversalUtil(WordprocessingMLPackage wordMLPackage) {
        if (wordMLPackage == null) {
            throw new NullPointerException();
        }
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart
                .getJaxbElement();
        Body body = wmlDocumentEl.getBody();
        // start recursive travel in document
        new TraversalUtil(body,
                new TraversalUtil.Callback() {
                    public List<Object> apply(Object o) {
                        if (o instanceof org.docx4j.wml.P) {

                            if(o.toString().startsWith("${")) {
                                Text text = new Text();
                                text.setValue("Welcome to Baeldun");
                                R replacement = new R();
                                replacement.getContent().add(text);
                                if(replacement != null) {
                                    ((P) o).getContent().clear();
                                    ((P) o).getContent().add(replacement);
                                }
                            }


                        }
                        return null;
                    }

                    public boolean shouldTraverse(Object o) {
                        return true;
                    }

                    public void walkJAXBElements(Object parent) {
                        List children = getChildren(parent);
                        if (children != null) {

                            for (Object o : children) {
                                o = XmlUtils.unwrap(o);
                                if (children.size() > 0) {
                                    System.out.println(children.get(0));
                                    System.out.println(children.get(0).getClass());
                                    if (children.get(0) instanceof org.docx4j.wml.P) {
                                        this.apply(o);
                                    }
                                }
                                if (this.shouldTraverse(o)) {
                                    walkJAXBElements(o);
                                }
                            }
                        }
                    }

                    public List<Object> getChildren(Object o) {
                        return TraversalUtil.getChildrenImpl(o);
                    }
                }
        );
        return wordMLPackage;
    }


    /*
    Innanzitutto, abbiamo creato il file che contiene l'immagine che vogliamo aggiungere nella nostra parte
    principale del documento, quindi abbiamo collegato l'array di byte che rappresenta l'immagine con l' oggetto
    wordMLPackage .
    Una volta creata la parte immagine, dobbiamo creare un oggetto Inline utilizzando il metodo createImageInline( ).
    Il metodo addImageToParagraph() incorpora l' oggetto Inline in un Drawing in modo
    che possa essere aggiunto a un'esecuzione.
    */
    public void workWithImage(WordprocessingMLPackage wordPackage, MainDocumentPart mainDocumentPart, String img) throws Exception {
        File image = new File(img);
        byte[] fileContent = Files.readAllBytes(image.toPath());
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage
                .createImagePart(wordPackage, fileContent);
        Inline inline = imagePart.createImageInline(
                "Baeldung Image (filename hint)", "Alt Text", 1, 2, false);
        P Imageparagraph = addImageToParagraph(inline);
        mainDocumentPart.getContent().add(Imageparagraph);
    }

    private static P addImageToParagraph(Inline inline) {
        ObjectFactory factory = new ObjectFactory();
        P p = factory.createP();
        R r = factory.createR();
        p.getContent().add(r);
        Drawing drawing = factory.createDrawing();
        r.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        return p;
    }


    public void createTable(WordprocessingMLPackage wordPackage) {
        int writableWidthTwips = wordPackage.getDocumentModel()
                .getSections().get(0).getPageDimensions().getWritableWidthTwips();
        int columnNumber = 3;
        Tbl tbl = TblFactory.createTable(3, 3, writableWidthTwips/columnNumber);
        List<Object> rows = tbl.getContent();
        for (Object row : rows) {
            Tr tr = (Tr) row;
            List<Object> cells = tr.getContent();
            for(Object cell : cells) {
                Tc td = (Tc) cell;
                Text text = new Text();
                text.setValue("Welcome to Baeldung");
                R r = new R();
                r.getContent().add(text);
                P p = new P();
                p.getContent().add(r);
                td.getContent().add(p);
            }
        }
    }


}
