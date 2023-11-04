package com.example.docx4j;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;

public interface DocxService {

    public WordprocessingMLPackage replaceTextWithTraversalUtil(WordprocessingMLPackage wordMLPackage);
    public void workWithImage(WordprocessingMLPackage wordPackage, MainDocumentPart mainDocumentPart, String img) throws Exception;

    public void createTable(WordprocessingMLPackage wordPackage);
}
