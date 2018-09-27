package org.zura.JournalFilter;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Properties;
import org.apache.commons.csv.CSVRecord;


public class Main {
    private static final String PROPERTIES_FILENAME = ".properties";
    public static void main(String[] args) {
        try {
            Properties prop = new Properties();
            prop.load(Files.newBufferedReader(Paths.get(PROPERTIES_FILENAME), StandardCharsets.UTF_8));
            if (args.length < 2) {
                System.out.println("入出力ファイル名を指定してください.");
                return;
            }
            String inFilename = args[0];
            String outFilename = args[1];

            Integer extPos = outFilename.lastIndexOf(".") + 1;
            String ext = outFilename.substring(extPos).toLowerCase();
            StoreType storeType = ext.equals("xlsx") ? StoreType.Xlsx : StoreType.Csv;
            String extText = storeType == StoreType.Xlsx ? "Excelファイル" : "CSVファイル";
            System.out.println("入力ファイル: " + inFilename);
            System.out.println("出力ファイル: " + outFilename + " (" + extText + ")");
            FilterManager c = new FilterManager(inFilename, outFilename, prop);

            c.run(storeType);
        } catch (FileNotFoundException e) {
            System.out.println(e);
        } catch (IOException e) {
            System.out.println(e);
        }
    }
}
