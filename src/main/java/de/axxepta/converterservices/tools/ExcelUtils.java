package de.axxepta.converterservices.tools;

import de.axxepta.converterservices.Const;
import de.axxepta.converterservices.utils.IOUtils;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.extractor.XSSFExportToXml;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;
import javax.xml.transform.TransformerException;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class ExcelUtils {

    private static Logger LOGGER = LoggerFactory.getLogger(ExcelUtils.class);

    private static final String XML_PROLOGUE    = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>"
            + System.lineSeparator();
    private static final String SHEET_INDENT    = "";
    private static final String ROW_INDENT      = "  ";
    private static final String COL_INDENT      = "    ";
    private static final String VALUE_INDENT    = "      ";
    private static final String FILE_EL         = "Workbook";
    public static final String SHEET_EL                = "Data";
    public static final String ROW_EL                  = "Row";
    public static final String COL_EL                  = "Cell";
    private static final String VALUE_EL        = "Value";
    public static final String DEF_SHEET_NAME          = "sheet0";
    public static final String DEF_ATT_SHEET           = "name";
    public static final String DEF_SEPARATOR           = ";";

    public static final String XML_SEPARATOR           = "\\r?\\n";


    public static List<String> fromTempExcel(String fileName, FileType type, boolean customXMLMapping, String sheetName,
                                             String separator, boolean indent, boolean columnFirst, boolean firstColName,
                                             boolean firstRowName, String fileEl, String sheetEl, String rowEl,
                                             String colEl, String attSheetName)
    {
        List<String> outputFiles = new ArrayList<>();
        try (FileInputStream file = new FileInputStream(Const.TEMP_FILE_PATH + "/" + fileName)) {
            outputFiles.addAll(
                    exportExcel(file, Const.TEMP_FILE_PATH, fileName, type, customXMLMapping, sheetName, separator, indent,
                    true, true, columnFirst, firstColName, firstRowName, fileEl, sheetEl, rowEl, colEl, attSheetName)
            );
        } catch (IOException ie) {
            LOGGER.error("Exception reading Excel file: " + ie.getMessage());
        }
        return outputFiles;
    }

    public static List<String> fromPipeExcel(String workPath, String inputFile, FileType type, boolean customXMLMapping, String sheetName,
                                             String separator, boolean indent, boolean wrapValue, boolean wrapSheet, boolean columnFirst, boolean firstColName,
                                             boolean firstRowName, String fileEl, String sheetEl, String rowEl,
                                             String colEl, String attSheetName)
            throws IOException
    {
        List<String> outputFiles;
        try (FileInputStream file = new FileInputStream(inputFile)) {
            outputFiles = exportExcel(file, workPath, IOUtils.filenameFromPath(inputFile), type, customXMLMapping, sheetName, separator,
                    indent, wrapValue, wrapSheet, columnFirst, firstColName, firstRowName, fileEl, sheetEl, rowEl, colEl, attSheetName).
                    stream().map(s -> IOUtils.pathCombine(workPath, s)).collect(Collectors.toList());
        }
        return outputFiles;
    }

    private static List<String> exportExcel(FileInputStream file, String path, String fileName, FileType type,
                                            boolean customXMLMapping, String sheetName, String separator, boolean indent,
                                            boolean wrapValue,  boolean wrapSheet, boolean columnFirst,
                                            boolean firstColName, boolean firstRowName, String fileEl, String sheetEl,
                                            String rowEl, String colEl, String attSheetName)
            throws IOException
    {
        List<String> outputFiles = new ArrayList<>();

            Workbook workbook = new XSSFWorkbook(file);
            if (type.equals(FileType.CSV)) {
                outputFiles.addAll(excelToCSV(fileName, workbook, sheetName, separator));
            }
            if (type.equals(FileType.XML)) {
                if (customXMLMapping)
                    outputFiles.addAll(excelCustomXMLMapping(path, fileName));
                else
                    outputFiles.add(excelToXML(path, fileName, workbook, sheetName, columnFirst, firstColName, firstRowName,
                            fileEl, sheetEl, rowEl, colEl, indent, wrapValue, wrapSheet, attSheetName));
            }
        return outputFiles;
    }

    private static List<String> excelToCSV(String fileName, Workbook workbook, String sheetName, String separator) {
        List<String> outputFiles = new ArrayList<>();
        List<Sheet> sheets = getSheets(workbook, sheetName);
        DataFormatter formatter = new DataFormatter(true);
        for (Sheet sheet : sheets) {
            String convertedFileName = CSVFileName(fileName, sheet.getSheetName());
            try (BufferedWriter writer = new BufferedWriter(new FileWriter(Const.TEMP_FILE_PATH + "/" + convertedFileName))) {
                int firstRow = sheet.getFirstRowNum();
                int lastRow = sheet.getLastRowNum();
                int firstColumn = Math.min(sheet.getRow(firstRow).getFirstCellNum(), sheet.getRow(firstRow + 1).getFirstCellNum());
                int lastColumn = Math.max(sheet.getRow(firstRow).getLastCellNum(), sheet.getRow(firstRow + 1).getLastCellNum());
                for (int rowNumber = firstRow; rowNumber < lastRow + 1; rowNumber++) {
                    Row row = sheet.getRow(rowNumber);
                    StringJoiner joiner = new StringJoiner(separator);
                    for (int colNumber = firstColumn; colNumber < lastColumn; colNumber++) {
                        Cell cell = row.getCell(colNumber);
                        joiner.add(formatter.formatCellValue(cell));
                    }
                    writer.write(joiner.toString() + System.lineSeparator());
                }
            } catch (IOException ie) {
                LOGGER.error("Exception writing to CSV file: " + ie.getMessage());
            }
            outputFiles.add(convertedFileName);
        }
        return outputFiles;
    }


    public static String excelSheetToHTMLString(String fileName, String sheetName, boolean firstRowHead) {
        StringBuilder builder = new StringBuilder();
        builder.append("<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\"" +
                " \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">")
                .append("<html xmlns=\"http://www.w3.org/1999/xhtml\">")
                .append("<head><title>").append(sheetName).append("</title></head>")
                .append("<body>");
        try (FileInputStream file = new FileInputStream(fileName)) {
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheet(sheetName);
            DataFormatter formatter = new DataFormatter(true);
            int firstRow = sheet.getFirstRowNum();
            int lastRow = sheet.getLastRowNum();
            int firstColumn = Math.min(sheet.getRow(firstRow).getFirstCellNum(), sheet.getRow(firstRow + 1).getFirstCellNum());
            int lastColumn = Math.max(sheet.getRow(firstRow).getLastCellNum(), sheet.getRow(firstRow + 1).getLastCellNum());

            builder.append("<table>");
            for (int rowNumber = firstRow; rowNumber < lastRow + 1; rowNumber++) {
                builder.append("<tr>");
                Row row = sheet.getRow(rowNumber);
                for (int colNumber = firstColumn; colNumber < lastColumn; colNumber++) {
                    builder.append((firstRowHead && (rowNumber == firstRow)) ? "<th>" : "<td>");
                    Cell cell = row.getCell(colNumber);
                    builder.append(xmlEncode(formatter.formatCellValue(cell)));
                    builder.append((firstRowHead && (colNumber == firstColumn)) ? "</th>" : "</td>");
                }
                builder.append("</tr>");
            }
            builder.append("</table>");
            builder.append("</body></html>");

        } catch (IOException ie) {
            LOGGER.error("Exception reading Excel file: " + ie.getMessage());
            builder.append("</body></html>");
        }
        return builder.toString();
    }

    private static List<String> excelCustomXMLMapping(String path, String fileName) throws IOException {
        List<String> customMappingFiles = new ArrayList<>();
        try {
            OPCPackage pkg = OPCPackage.open(fileName);
            XSSFWorkbook wb = new XSSFWorkbook(pkg);
            for (XSSFMap map : wb.getCustomXMLMappings()) {
                XSSFExportToXml exporter = new XSSFExportToXml(map);
                String outputFile = XMLMappingFileName(fileName, map.hashCode());
                try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
                    exporter.exportToXML(baos, true);
                    try (OutputStream outputStream = new FileOutputStream(path + "/" + outputFile)) {
                        baos.writeTo(outputStream);
                        customMappingFiles.add(outputFile);
                    }
                }
            }
            pkg.close();
        } catch (InvalidFormatException | SAXException | TransformerException ife) {
            LOGGER.error("Exception writing to CSV file: " + ife.getMessage());
        }
        return customMappingFiles;
    }


    public static String CSVToExcel(String fileName, String sheetName, String separator) {
        String outputFile = XLSXFileName(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook();
        try {
            FileOutputStream out = new FileOutputStream(new File(Const.TEMP_FILE_PATH + "/" + outputFile));
            XSSFSheet sheet = workbook.createSheet(sheetName);
            try (FileInputStream fis = new FileInputStream(Const.TEMP_FILE_PATH + "/" + fileName)) {
                Scanner scanner = new Scanner(fis);
                int rowId = 0;
                while (scanner.hasNext()) {
                    String line = scanner.nextLine();
                    String[] cellContents = line.split(separator);
                    XSSFRow row = sheet.createRow(rowId++);
                    int cellId = 0;
                    for (String el : cellContents)
                    {
                        Cell cell = row.createCell(cellId++);
                        cell.setCellValue(el);
                    }
                }
            }
            workbook.write(out);
            out.close();
        } catch (IOException ie) {
            LOGGER.error("Exception writing to XLSX file: " + ie.getMessage());
        }
        return outputFile;
    }

    public static String serviceXMLToExcel(String fileName, String sheetName, String row, String col, String separator, String... outputFileName) {
        String outputFile = Const.TEMP_FILE_PATH + "/" + (outputFileName.length > 0 ? outputFileName[0]: XLSXFileName(fileName));
        String type = Const.DATA_TYPE_ATT;
        boolean cellFormat = true;
        return XMLToExcel(fileName, false, sheetName, row, col, type, cellFormat, separator, outputFile);
    }

    public static String XMLToExcel(String fileName, boolean multiSheet, String sheetName, String row, String col, String type,
                                    boolean cellFormat, String separator, String outputFile) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        try {

            ExcelContentHandler handler = new ExcelContentHandler(workbook, multiSheet, sheetName, row, col, type, cellFormat, separator);
            File file = new File(fileName);
            SAXParserFactory factory = SAXParserFactory.newInstance();
            SAXParser parser = factory.newSAXParser();
            parser.parse(file, handler);

            // auto-adjust columns widths
            if (multiSheet) {
                IntStream.range(0, workbook.getNumberOfSheets()).forEach(workbook::getSheetAt);
            } else {
                adjustColumns(workbook.getSheet(sheetName));
            }

            try (FileOutputStream out = new FileOutputStream(new File(outputFile))) {
                workbook.write(out);
            }

        } catch (SAXException | ParserConfigurationException sx) {
            LOGGER.error("Exception reading " + sx.getMessage());
        } catch (IOException ie) {
            LOGGER.error("Exception writing to XLSX file: " + ie.getMessage());
        }
        return outputFile;
    }

    private static void adjustColumns(Sheet sheet) {
        int lastRow = sheet.getLastRowNum();
        int lastColumn = 0;
        for (int r = 0; r <= lastRow; r++) {
            lastColumn = Math.max(lastColumn, sheet.getRow(r).getLastCellNum());
        }
        for (int c = 0; c <= lastColumn; c++) {
            sheet.autoSizeColumn((short) c);
        }
    }

    private static String excelToXML(String path, String fileName, Workbook workbook, String sheetName, boolean columnFirst,
                                   boolean firstColName, boolean firstRowName,
                                   String fileEl, String sheetEl, String rowEl, String colEl,
                                   boolean indent, boolean wrapValue, boolean wrapSheet, String attSheetName) {
        List<Sheet> sheets = getSheets(workbook, sheetName);
        FormulaEvaluator evaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
        DataFormatter formatter = new DataFormatter(true);

        String outputFile = XMLFileName(fileName);
        String file_indent = "";
        String columnElement = colEl.equals("") ? COL_EL : colEl;
        String rowElement = rowEl.equals("") ? ROW_EL : rowEl;
        try (BufferedWriter writer =
                     new BufferedWriter(new OutputStreamWriter(new FileOutputStream(path + "/" + outputFile), StandardCharsets.UTF_8)))
        {
            writer.write(XML_PROLOGUE);
            if (wrapSheet) {
                writeTag(writer, TagType.open, fileEl.equals("") ? FILE_EL : fileEl, indent, false, "");
                file_indent = "  ";
            }
            for (Sheet sheet : sheets) {
                if (attSheetName.equals("")) {
                    writeTag(writer, TagType.open, sheetEl.equals("") ? SHEET_EL : sheetEl, indent,false, SHEET_INDENT + file_indent);
                } else {
                    writeTag(writer, TagType.open, sheetEl.equals("") ? SHEET_EL : sheetEl, indent, false, SHEET_INDENT + file_indent,
                            attSheetName, sheet.getSheetName());
                }
                int firstRow = sheet.getFirstRowNum();
                int lastRow = sheet.getLastRowNum();
                int firstColumn = Math.min(sheet.getRow(firstRow).getFirstCellNum(), sheet.getRow(firstRow + 1).getFirstCellNum());
                int lastColumn = Math.max(sheet.getRow(firstRow).getLastCellNum(), sheet.getRow(firstRow + 1).getLastCellNum());
                if (columnFirst) {
                    //
                } else {
                    for (int rowNumber = firstRow; rowNumber < lastRow + 1; rowNumber++) {
                        Row row = sheet.getRow(rowNumber);
                        if (row == null) {
                            LOGGER.warn("EXCEL CONVERSION ERROR - ROW NULL IN ROW " + rowNumber);
                            continue;
                        }
                        writeTag(writer, TagType.open, rowElement, indent, false, ROW_INDENT + file_indent,
                                "RowNumber", Integer.toString(row.getRowNum()));
                        for (int colNumber = firstColumn; colNumber < lastColumn; colNumber++) {
                            Cell cell = row.getCell(colNumber);
                            String cellType = (cell != null) ? getCellType(cell) : "_";
                            String cellValue = xmlEncode( cellType.equals("text") ?
                                    ((XSSFCell) cell).getRichStringCellValue().toString() :
                                    formatter.formatCellValue(cell, evaluator) );
                            if (wrapValue) {
                                writeTag(writer, TagType.open, columnElement, indent, false,
                                        COL_INDENT + file_indent,
                                        "ref", (cell != null) ? cell.getAddress().toString() : "_",
                                        "column-number", (cell != null) ? Integer.toString(cell.getColumnIndex()) : "_",
                                        "data-type", cellType);
                                writeElement(writer, VALUE_EL,
                                        cellValue, indent, VALUE_INDENT + file_indent);
                                writeTag(writer, TagType.close, columnElement, indent, false, COL_INDENT + file_indent);
                            } else {
                                writeElement(writer, columnElement,
                                        cellValue, indent, COL_INDENT + file_indent,
                                        "ref", (cell != null) ? cell.getAddress().toString() : "_",
                                        "column-number", (cell != null) ? Integer.toString(cell.getColumnIndex()) : "_",
                                        Const.DATA_TYPE_ATT, cellType);
                            }
                        }
                        writeTag(writer, TagType.close, rowElement, indent, false, ROW_INDENT + file_indent);
                    }
                }
                writeTag(writer, TagType.close, sheetEl.equals("") ? SHEET_EL : sheetEl, indent, false, SHEET_INDENT + file_indent);
            }
            if (wrapSheet) {
                writeTag(writer, TagType.close, fileEl.equals("") ? FILE_EL : fileEl, indent, false, "");
            }
        } catch (IOException ie){
            LOGGER.error("Exception writing to XML file: " + ie.getMessage());
        }
        return outputFile;
    }

    private static String xmlEncode(String val) {
        return val.replaceAll("&([^;\\W]*([^;\\w]|$))", "&amp;$1").
                replaceAll("<","&lt;").replaceAll(">", "&gt;");
    }

    private static String getCellType(Cell cell) {
        String typeString = cell.getCellType().toString().toLowerCase();
        switch (typeString) {
            case "string": return "text";
            case "numeric": if (DateUtil.isCellDateFormatted(cell)) {
                return "date";
            } else {
                return "number";
            }
            case "formula": return "number";
            default: return typeString;
        }
    }

    private static List<Sheet> getSheets(Workbook workbook, String sheetName) {
        List<Sheet> sheets = new ArrayList<>();
        if (!sheetName.equals("") && (workbook.getSheet(sheetName) != null)) {
            sheets.add(workbook.getSheet(sheetName));
        } else {
            for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
                sheets.add(workbook.getSheetAt(s));
            }
        }
        return sheets;
    }

    public static ByteArrayOutputStream XMLToCSV(String fileName, String rowTag, String columnTag, String delimiter)
            throws ParserConfigurationException, SAXException, IOException
    {
        CSVContentHandler handler = new CSVContentHandler(rowTag, columnTag, delimiter);
        File file = new File(fileName);
        SAXParserFactory factory = SAXParserFactory.newInstance();
        SAXParser parser = factory.newSAXParser();
        parser.parse(file, handler);
        ByteArrayOutputStream builderStream = handler.getOutputStream();
        return builderStream;
    }

    private static void writeTag(BufferedWriter writer, TagType tag, String name, boolean indent, boolean tagWithContent,
                                 String indentString, String ... attributes) throws IOException {
        StringBuilder builder = new StringBuilder();
        builder.append((indent && (!tagWithContent || tag.equals(TagType.open))) ? indentString : "");
        if (tag.equals(TagType.open)) {
            builder.append("<").append(name);
            for (int i = 0; i < attributes.length / 2; i++) {
                builder.append(" ").append(attributes[2 * i]).append("=\"").append(attributes[2 * i + 1]).append("\"");
            }
            builder.append(">");
        } else {
            builder.append("</").append(name).append(">");
        }
        builder.append((indent && (!tagWithContent || tag.equals(TagType.close))) ? System.lineSeparator() : "");
        writer.write(builder.toString());
    }

    private static void writeElement(BufferedWriter writer, String name, String content,
                                 boolean indent, String indentTagString,
                                     String ... attributes) throws IOException {
        writeTag(writer, TagType.open, name, indent, true, indentTagString, attributes);
        writer.write(content);
        writeTag(writer, TagType.close, name, indent, true, indentTagString);
    }

    private static String XLSXFileName(String name) {
        return name.substring(0, name.lastIndexOf(".")) + ".xlsx";
    }

    private static String XMLFileName(String xlsName) {
        return xlsName.substring(0, xlsName.lastIndexOf(".")) + ".xml";
    }

    private static String XMLMappingFileName(String xlsName, int hash) {
        return xlsName.substring(0, xlsName.lastIndexOf(".")) + "_" + Integer.toString(hash) + ".xml";
    }

    private static String CSVFileName(String xlsName, String sheetName) {
        return xlsName.substring(0, xlsName.lastIndexOf(".")) + "_" + sheetName + ".csv";
    }


    public enum FileType {
        XLSX,
        CSV,
        XML
    }

    private enum TagType {
        open,
        close
    }

    public static String excelSheetTransformString(String fileName, String sheetName, String root,
                                                   String nlSeparatedMappingList, boolean firstRowHead) {
        List<String[]> mapping = listToLinkedMap(nlSeparatedMappingList);
        XMLOutputFactory factory = XMLOutputFactory.newInstance();
        StringWriter result = new StringWriter();
        try(FileInputStream file = new FileInputStream(fileName)) {
            XMLStreamWriter writer = factory.createXMLStreamWriter(result);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheet(sheetName);
            DataFormatter formatter = new DataFormatter(true);
            int firstRow = sheet.getFirstRowNum();
            int lastRow = sheet.getLastRowNum();
            int firstColumn = Math.min(sheet.getRow(firstRow).getFirstCellNum(),
                    sheet.getRow(firstRow + 1).getFirstCellNum());
            int lastColumn = Math.max(sheet.getRow(firstRow).getLastCellNum(),
                    sheet.getRow(firstRow + 1).getLastCellNum());

            List<String> headers = new ArrayList<>();
            Row row1 = sheet.getRow(firstRow);
            for (int colNumber = firstColumn; colNumber < lastColumn; colNumber++) {
                headers.add(firstRowHead ? formatter.formatCellValue(row1.getCell(colNumber)):
                        /*org.apache.poi.ss.util.*/CellReference.convertNumToColString(row1.getCell(colNumber).getColumnIndex()));
            }

            writer.writeStartDocument();
            writer.writeStartElement(root);
            for (int rowNumber = firstRow + (firstRowHead ? 1 : 0); rowNumber < lastRow + 1; rowNumber++) {
                Row row = sheet.getRow(rowNumber);

                int[] openElements = {0};
                String[] lastElement = {""};
                mapping.forEach(k -> {
                    int col = headers.indexOf(k[0]) + firstColumn;
                    if (col > firstColumn - 1) {
                        try {
                            Cell cell = row.getCell(col);
                            String val = formatter.formatCellValue(cell);
                            String[][] diff = pathDiff(lastElement[0], k[1]);
                            closeWriterElements(writer, diff[0].length - (lastElement[0].contains("@") ? 1 :0));
                            for (int e = 0; e < diff[1].length; e++) {
                                if (diff[1][e].startsWith("@")) {
                                    writer.writeAttribute(diff[1][e].substring(1), val);
                                } else {
                                    writer.writeStartElement(diff[1][e]);
                                    if (e == diff[1].length - 1)
                                        writer.writeCharacters(val);
                                }
                            }
                            if (diff[1].length == 0)
                                writer.writeCharacters(val);
                            lastElement[0] = k[1];
                            String[] lastPath = k[1].split("/");
                            openElements[0] = lastPath.length - (lastPath[lastPath.length - 1].startsWith("@") ? 1 : 0);
                        } catch (XMLStreamException | IllegalArgumentException xs) {}
                    }
                });

                closeWriterElements(writer, openElements[0]);
            }
            writer.writeEndElement();
            writer.writeEndDocument();
            writer.close();

        } catch (XMLStreamException|IOException xe) {
            return "<xml-transformation-error>" + xe.getMessage(). replaceAll("&", "&amp;").
                    replaceAll("<", "&lt;").replaceAll(">", "&gt;") +
                    "<xml-transformation-error>";
        }
        return result.toString();
    }


    private static String[][] pathDiff(String last, String current) throws IllegalArgumentException {
        if (last.equals(current))
            throw new IllegalArgumentException("Consecutive path definitions must not be equal.");
        String[][] diff = {new String[0], new String[0]};
        String[] old = (last.equals("")) ? new String[0] : last.split("/");
        String[] neww = current.split("/");
        int minLength = Math.min(old.length, neww.length);
        if (old.length == 0)
            diff[1] = neww;
        for (int i = 0; i < minLength; i++) {
            if (!old[i].equals(neww[i])) {
                diff[0] = Arrays.copyOfRange(old, i, old.length);
                diff[1] = Arrays.copyOfRange(neww, i, neww.length);
                break;
            }
            if (i + 1 == minLength) {
                diff[0] = (old.length > minLength) ?
                        Arrays.copyOfRange(old, minLength, old.length) : new String[0];
                diff[1] = (neww.length > minLength) ?
                        Arrays.copyOfRange(neww, minLength, neww.length) : new String[0];
            }
        }
        return diff;
    }


    private static void closeWriterElements(XMLStreamWriter writer, int num) throws XMLStreamException {
        for (int i = 0; i < num; i++) writer.writeEndElement();
    }

    private static List<String[]> listToLinkedMap(String nlMappingList) {
        String[] mappingList = nlMappingList.split("\\r?\\n");
        List<String[]> map = new LinkedList<>();
        for (String mapLine : mappingList) {
            if (!mapLine.equals("")) {
                String[] entry = mapLine.split("\\t");
                map.add(new String[]{entry[0], entry[1]});
            }
        }
        return map;
    }


}
