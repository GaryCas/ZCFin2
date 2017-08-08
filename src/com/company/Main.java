package com.company;

public class Main {

    public static void main(String[] args) {
	// write your code here
        try {
            Copy.copyHeaderAndFooter(System.getProperty("user.dir") + "\\template.docx",
                    System.getProperty("user.dir") + "\\target.docx");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
