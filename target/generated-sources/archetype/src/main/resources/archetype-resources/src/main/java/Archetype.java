#set( $symbol_pound = '#' )
#set( $symbol_dollar = '$' )
#set( $symbol_escape = '\' )

//package org.docx4j.samples;


import java.util.HashMap;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Document;

public class Archetype {

    public static void main(String[] args) throws Exception {

        String inputfilepath = System.getProperty("user.dir") + "/src/main/resources/inputFile.docx";
        boolean save = true;
        String outputfilepath = System.getProperty("user.dir")
                + "/src/main/resources/output.docx";

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                .load(new java.io.File(inputfilepath));
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        String xml = XmlUtils.marshaltoString(documentPart.getJaxbElement(), true);
        HashMap<String, String> mappings = new HashMap<String, String>();
        mappings.put("name", "world");

        Object obj = XmlUtils.unmarshallFromTemplate(xml, mappings);

        documentPart.setJaxbElement((Document) obj);

        // Save it
        if (save) {
            SaveToZipFile saver = new SaveToZipFile(wordMLPackage);
            saver.save(outputfilepath);
        } else {
            System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true,
                    true));
        }
    }

}