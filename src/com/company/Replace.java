package com.company;

import org.docx4j.openpackaging.Base;
import org.docx4j.openpackaging.contenttype.ContentTypeManager;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io.Load;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

/**
 * Created by Desktop on 8/3/2017.
 */
public class Replace {

    public static void replaceHeader(String file, String baseDir, String expandedDir) throws Exception {
        WordprocessingMLPackage wordMLPackage = null;
        if (file.endsWith(".docx")) {
            try {
                wordMLPackage = WordprocessingMLPackage.load(new File(baseDir + file));

                InputStream in = new FileInputStream(baseDir + expandedDir + "/[Content_Types].xml");
                ContentTypeManager externalCtm = new ContentTypeManager();

                externalCtm.parseContentTypesFile(in);

                // Example of a part which become a rel of the word document
                in = new FileInputStream(baseDir + expandedDir + "/word/header1.xml");
                attachForeignPart(wordMLPackage,
                        wordMLPackage.getMainDocumentPart(),
                        externalCtm,
                        "word/header1.xml",
                        in);

                // Example of a part which become a rel of the package
                in = new FileInputStream(baseDir + expandedDir + "/docProps/app.xml");
                attachForeignPart(wordMLPackage, wordMLPackage,
                        externalCtm, "docProps/app.xml", in);


                wordMLPackage.save(new File(baseDir + "header-added.docx"));
            } catch (Docx4JException e) {
                e.printStackTrace();
            }
        }
    }

    public static void attachForeignPart(WordprocessingMLPackage wordMLPackage,
                                         Base attachmentPoint,
                                         ContentTypeManager foreignCtm,
                                         String resolvedPartUri, InputStream is) throws Exception {
        Part foreignPart = Load.getRawPart(is, foreignCtm, resolvedPartUri, null);
        // the null means this won't work for an AlternativeFormatInputPart
        attachmentPoint.addTargetPart(foreignPart);
        // Add content type
        ContentTypeManager packageCtm = wordMLPackage.getContentTypeManager();
        packageCtm.addOverrideContentType(foreignPart.getPartName().getURI(), foreignPart.getContentType());

    }
}
