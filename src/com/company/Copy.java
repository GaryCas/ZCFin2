package com.company;

import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import java.io.File;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Created by Desktop on 8/3/2017.
 */
public class Copy {
    static ObjectFactory factory = Context.getWmlObjectFactory();

    public static void copyHeaderAndFooter(String sourcefile, String targetfile) throws Exception {
        WordprocessingMLPackage source;
        WordprocessingMLPackage target;
        if (sourcefile.endsWith(".docx")) {
            try {
                source = WordprocessingMLPackage.load(new File(sourcefile));
                target = WordprocessingMLPackage.createPackage();

                List<HeaderPart> headerParts = extractHeaderParts(source);

                for (SectionWrapper sectionWrapper : target.getDocumentModel().getSections()) {
                    sectionWrapper.getSectPr();
                }

                for (SectionWrapper sectionWrapper : source.getDocumentModel().getSections()) {
                    sectionWrapper.getSectPr();
                }


                SectPr sectPr = removeExistingReferences(target);


                // Find the part we want to copy
                RelationshipsPart rp = source.getMainDocumentPart().getRelationshipsPart();
                Relationship relationship = null;

                ArrayList<HeaderReference> headerReferences = getHeaderReferences(rp);

                for (HeaderPart headerPart : headerParts) {

                    relationship = target.getMainDocumentPart().addTargetPart(headerPart);
                    createAndAddHeader(relationship, sectPr);
                }


//                rp = source.getMainDocumentPart().getRelationshipsPart();
//                rel = rp.getRelationshipByType(Namespaces.FOOTER);
//                p = rp.getPart(rel);
//
//                p.setPartName(new PartName("/word/footer1.xml"));
//
//                // Now try adding it
//                target.getMainDocumentPart().addTargetPart(p);

                target.save(new File(targetfile));
            } catch (Docx4JException e) {
                e.printStackTrace();
            }
        }

    }

    private static ArrayList<HeaderReference> getHeaderReferences(RelationshipsPart rp) {
        ArrayList<Object> headerReferences = new ArrayList();

        HeaderReference headerReference = new HeaderReference();

        for (Relationship relationship : rp.getRelationships().getRelationship()) {
            if (relationship.getType().contains("header")){
                headerReferences.add((Object) relationship);
            }
        }

        return null;
    }

    private static List<HeaderPart> extractHeaderParts(WordprocessingMLPackage source) {
        ArrayList<HeaderPart> headerParts = new ArrayList<HeaderPart>();
        for (PartName partName : source.getParts().getParts().keySet()) {
            if(partName.getName().contains("header")){
                headerParts.add((HeaderPart) source.getParts().get(partName));
            }
        }

        return headerParts;
    }

    private static SectPr removeExistingReferences(WordprocessingMLPackage wordMLPackage) {
        List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();

        SectPr sectionProperties = sections.get(sections.size() - 1).getSectPr();

        if (sectionProperties == null) {
            sectionProperties = factory.createSectPr();
            wordMLPackage.getMainDocumentPart().addObject(sectionProperties);
            sections.get(sections.size() - 1).setSectPr(sectionProperties);
        }

         /*
          * Remove Header if it already exists.
          */
        List<CTRel> relations = sectionProperties.getEGHdrFtrReferences();
        Iterator<CTRel> relationsItr = relations.iterator();
        while (relationsItr.hasNext()) {
            CTRel relation = relationsItr.next();
            if (relation instanceof HeaderReference) {
                relationsItr.remove();
            }
        }

        return sectionProperties;
    }

    private static void createAndAddHeader(Relationship relationship, SectPr sectionProperties) {
        // default header
        for (CTRel ctRel : sectionProperties.getEGHdrFtrReferences()) {

        }

    }
}
