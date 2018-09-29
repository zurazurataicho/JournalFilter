package org.zura.JournalFilter;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.Exception;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;


public class FilterManager {
	private final String CSV_CHARSET;// = "MS932";	// CSVファイルはShift_JIS(cp932/MS932)
	private final String REGEX_OFFICE_FILE_EXTENSIONS;
    private final String DEFAULT_REGEX_EVENTINFO;
    private Integer beforeFilterd = 0;
    private Integer afterFilterd = 0;
    private BufferedReader csvIn;
    private String outFilename;
    private Properties prop;
    private List<String> columnHeaders = new ArrayList<>();
    private List<String> outColumnHeaders = new ArrayList<>();
    FilterManager(String inFilename, String outFilename, Properties prop) throws IOException {
        CSV_CHARSET = prop.getProperty("CSV_CHARSET");
        REGEX_OFFICE_FILE_EXTENSIONS = prop.getProperty("DEFAULT_REGEX_OFFICE_FILE_EXTENSIONS");
        DEFAULT_REGEX_EVENTINFO = prop.getProperty("DEFAULT_REGEX_EVENTINFO");
        csvIn = Files.newBufferedReader(Paths.get(inFilename), Charset.forName(CSV_CHARSET));
        this.outFilename = outFilename;
        this.prop = prop;
        setHeaders();
        setOutputHeaders();
    }
    private void setHeaders() {
        columnHeaders.add(prop.getProperty("COL_TEXT_TIMESTAMP"));
        columnHeaders.add(prop.getProperty("COL_TEXT_USN"));
        columnHeaders.add(prop.getProperty("COL_TEXT_FILENAME"));
        columnHeaders.add(prop.getProperty("COL_TEXT_FULLPATH"));
        columnHeaders.add(prop.getProperty("COL_TEXT_EVENTINFO"));
        columnHeaders.add(prop.getProperty("COL_TEXT_SOURCEINFO"));
        columnHeaders.add(prop.getProperty("COL_TEXT_FILEATTRIBUTE"));
    }
    private void setOutputHeaders() {
        outColumnHeaders.add(prop.getProperty("OUTPUT_COL_TEXT_NO"));
        outColumnHeaders.add(prop.getProperty("OUTPUT_COL_TEXT_TIMESTAMP"));
        outColumnHeaders.add(prop.getProperty("OUTPUT_COL_TEXT_FILENAME"));
        outColumnHeaders.add(prop.getProperty("OUTPUT_COL_TEXT_FULLPATH"));
        outColumnHeaders.add(prop.getProperty("OUTPUT_COL_TEXT_EVENTINFO"));
        outColumnHeaders.add(prop.getProperty("OUTPUT_COL_TEXT_FILEATTRIBUTE"));
    }
    private List<CSVRecord> read() throws IOException {
        CSVParser parser = CSVFormat
            .RFC4180
			.withHeader()
            .withIgnoreEmptyLines(true)
            .withIgnoreSurroundingSpaces(true)
            .parse(csvIn);
        return parser.getRecords();
    }
    private void filter(IRowStore store) throws IOException {
        beforeFilterd = 0;
        afterFilterd = 0;
        for (CSVRecord record : read()) {
            String fileName = record.get(columnHeaders.get(2));
            beforeFilterd += 1;
			Pattern p = Pattern.compile(REGEX_OFFICE_FILE_EXTENSIONS, Pattern.CASE_INSENSITIVE);
			Matcher m = p.matcher(fileName);
            if (!m.find()) {
                // オフィスファイルではなかったらスキップ
				continue;
			}
            String eventInfo = record.get(columnHeaders.get(4));
			p = Pattern.compile(DEFAULT_REGEX_EVENTINFO, Pattern.CASE_INSENSITIVE);
			m = p.matcher(eventInfo);
            if (!m.find()) {
                // 対象イベントではなかった
				continue;
			}

            String timeStamp = record.get(columnHeaders.get(0));
            String usn = record.get(columnHeaders.get(1));
            String fullPath = record.get(columnHeaders.get(3));
            String sourceInfo = record.get(columnHeaders.get(5));
            String fileAttr = record.get(columnHeaders.get(6));
            afterFilterd += 1;
			store.storeRow(afterFilterd, timeStamp, fileName, fullPath, eventInfo, fileAttr);
        }
		store.close();
    }
    private void removePrivacyInformation() throws IOException {
        InputStream poiIs = new FileInputStream(outFilename);
        POIFSFileSystem poiFs = new POIFSFileSystem(poiIs);
        poiIs.close();
        DirectoryEntry poiDir = poiFs.getRoot();
        SummaryInformation info = PropertySetFactory.newSummaryInformation();
        info.removeAuthor();
        info.removeLastAuthor();








    }
    public void run(StoreType type) throws IOException {
        OutputStream outStream = Files.newOutputStream(Paths.get(outFilename));
        switch (type) {
        case Csv:
            filter(new CsvStore(outStream, outColumnHeaders, prop));
            break;
        case Xlsx:
            Integer pos = outFilename.lastIndexOf("/") + 1;
            String outFilenameForHeader = outFilename.substring(pos);
            filter(new XlsxStore(outStream, outColumnHeaders, prop, Arrays.asList(outFilenameForHeader, prop.getProperty("XLSX_SHEET_NAME"))));
            removePrivacyInformation();
            break;
        }
        outStream.flush();
		outStream.close();
        System.out.println("フィルタ前: " + beforeFilterd + "行");
        System.out.println("フィルタ後: " + afterFilterd + "行");
    }
}
