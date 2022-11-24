package io.operon.component.integration.writeexcel;

import java.util.Collections;
import java.util.Collection;
import java.util.Map;
import java.util.HashMap;
import java.util.stream.Stream;
import java.util.Set;
import java.util.List;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.io.IOException;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.BorderStyle;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.ss.usermodel.Picture;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import io.operon.runner.model.exception.*;
import io.operon.runner.statement.Statement;
import io.operon.runner.node.AbstractNode;
import io.operon.runner.node.Node;
import io.operon.runner.node.type.*;
import io.operon.runner.processor.function.core.raw.RawToStringType;
import io.operon.runner.processor.function.core.date.DateNow;
import io.operon.runner.system.IntegrationComponent;
import io.operon.runner.system.integration.BaseComponent;
import io.operon.runner.util.JsonUtil;


import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;

//
// This component is a "producer", i.e. it only writes data.
//
public class WriteExcelIntegrationComponent extends BaseComponent implements IntegrationComponent {

    private static Logger log = LogManager.getLogger(WriteExcelIntegrationComponent.class);

    public WriteExcelIntegrationComponent() {
        log.debug("WriteExcelIntegrationComponent :: constructor");
    }
    
    public OperonValue produce(OperonValue currentValue) throws OperonComponentException {
        log.debug("WriteExcelIntegrationComponent :: produce");
        try {
            Info info = resolve(currentValue);
            
            if (info.fileName == null) {
                throw new Exception("Missing fileName");
            }
            
            OperonValue result = this.handleTask(currentValue, info);
            return result;
        } catch (Exception e) {
            OperonComponentException oce = new OperonComponentException(e.getMessage());
            throw oce;
        }
    }

    private void debug(Info info, String msg) {
        if (info.debug) {
            System.out.println(msg);
        }
    }

    private OperonValue handleTask(OperonValue currentValue, Info info) throws Exception {
        Statement stmt = currentValue.getStatement();
        
        XSSFWorkbook wb = null;
        XSSFSheet sheet = null;
        
        //System.out.println("-------------------");
        
        // Map of <styleName, CellSTyle>
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
        
        if (info.load == false) {
            debug(info, "== handleTask :: overwrite or create new excel");
            wb = new XSSFWorkbook();
            sheet = wb.createSheet("sheet");
            if (info.columnWidth != null) {
                sheet.setDefaultColumnWidth(info.columnWidth);
            }
            if (info.landscape) {
                sheet.getPrintSetup().setLandscape(true);
            }
        }
        
        else {
            debug(info, "== handleTask :: load existing excel");
            File file = new File(info.fileName);
            FileInputStream fis = new FileInputStream(file);
            wb = new XSSFWorkbook(fis);
            sheet = wb.getSheet("sheet");
            if (info.columnWidth != null) {
                sheet.setDefaultColumnWidth(info.columnWidth);
            }
            if (info.landscape) {
                sheet.getPrintSetup().setLandscape(true);
            }
        }
        
        //
        // Process data:
        //
        for (ExcelWriteCommand cmd : info.commands) {
            debug(info, "CMD :: " + cmd);
            XSSFRow row = sheet.getRow(cmd.row());
            if (row == null) {
                row = sheet.createRow(cmd.row());
            }
            XSSFCell cell = row.getCell(cmd.cell());
            if (cell == null) {
                cell = row.createCell(cmd.cell());
            }
            if (cmd.styleName != null) {
                CellStyle cellStyle = styles.get(cmd.styleName());
                if (cellStyle == null) {
                    cellStyle = wb.createCellStyle();
                    if (cmd.fill() != null) {
                        cellStyle.setFillForegroundColor(cmd.fill);
                        //cellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    }
                    if (cmd.bold()) {
                        Font font = wb.createFont();
                        //font.setFontName("Courier New");
                        font.setBold(true);
                        cellStyle.setFont(font);
                    }
                    if (cmd.wrap()) {
                        cellStyle.setWrapText(true);
                    }
                    else if (cmd.wrap() == false) {
                        cellStyle.setWrapText(false);
                    }
                    if (cmd.borderBottom()) {
                        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
                    }
                    if (cmd.borderTop()) {
                        cellStyle.setBorderTop(BorderStyle.MEDIUM);
                    }
                    if (cmd.borderLeft()) {
                        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
                    }
                    if (cmd.borderRight()) {
                        cellStyle.setBorderRight(BorderStyle.MEDIUM);
                    }
                    //System.out.println("1) wrap: " + cmd.wrap() + ", cellStyleWrap=" + cellStyle.getWrapText() + ", style=" + cmd.styleName());
                    //
                    // NOTE: re-using the same CellStyle causes that if style is changed, then the referenced style will get into wrong state,
                    //         even AFTER it should have been rendered on node. Therefore we must deep-clone the style.
                    CellStyle cellStyleClone = wb.createCellStyle();
                    cellStyleClone.cloneStyleFrom(cellStyle);
                    styles.put(cmd.styleName(), cellStyleClone);
                }
                else {

                    //System.out.println("2) wrap: " + cmd.wrap() + ", cellStyleWrap=" + cellStyle.getWrapText() + ", style=" + cmd.styleName());
                    //
                    // NOTE: overwrite the cellStyle -styles, e.g. "wrap: false" was not set in the first instance, but we want to augment the style
                    // TODO: create own method to augment the recorded style.
                    if (cmd.wrap()) {
                        cellStyle.setWrapText(true);
                    }
                    else if (cmd.wrap() == false) {
                        cellStyle.setWrapText(false);
                    }
                }
                //System.out.println(cmd);
                cell.setCellStyle(cellStyle);
            }
            if (cmd.cellValue() != null && cmd.cellValue() instanceof EmptyType == false) {
                if (cmd.cellValue() instanceof StringType) {
                    cell.setCellValue(((StringType) cmd.cellValue()).getJavaStringValue()); // This method returns nothing. 
                }
                else {
                    cell.setCellValue(cmd.cellValue().toString());
                }
            }
            else {
                //System.out.println("Skipping cell-value");
            }
        }
        
        //
        // Insert image(s)
        //
        if (info.imageName != null) {
            insertImage(wb, sheet, info);
        }
        
        // 
        // Output Excel:
        //
        FileOutputStream fos = new FileOutputStream(info.fileName);
        wb.write(fos);
        fos.close();
        return new EmptyType(stmt);
    }

    // 
    // @param imageFormat "png"
    //
    private static void insertImage(XSSFWorkbook wb, XSSFSheet sheet, Info info) throws Exception {
        BufferedImage image = ImageIO.read(new File(info.imageName));
        ByteArrayOutputStream baps = new ByteArrayOutputStream();
        ImageIO.write(image, info.imageFormat, baps);

        int pictureIdx = wb.addPicture(baps.toByteArray(), Workbook.PICTURE_TYPE_PNG);

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFCreationHelper helper = wb.getCreationHelper();
        XSSFClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(info.imageXPos);
        anchor.setRow1(info.imageYPos);

        Picture picture = drawing.createPicture(anchor, pictureIdx);
        if (info.imageScale != null) {
            picture.resize(info.imageScale);
        }
        else {
            picture.resize();
        }
    }

    public Info resolve(OperonValue currentValue) throws Exception {
        OperonValue currentValueCopy = currentValue;
        
        ObjectType jsonConfiguration = this.getJsonConfiguration();
        jsonConfiguration.getStatement().setCurrentValue(currentValueCopy);
        List<PairType> jsonPairs = jsonConfiguration.getPairs();
        
        Info info = new Info();
        
        for (PairType pair : jsonPairs) {
            String key = pair.getKey();
            //System.out.println("for :: " + key);
            OperonValue currentValueCopy2 = currentValue;
            pair.getStatement().setCurrentValue(currentValueCopy);
            switch (key.toLowerCase()) {
                case "\"filename\"":
                    String fileNameVal = ((StringType) pair.getValue().evaluate()).getJavaStringValue();
                    info.fileName = fileNameVal;
                    break;
                case "\"commands\"":
                    List<Node> commandsNodes = ((ArrayType) pair.getValue().evaluate()).getValues();
                    List<ExcelWriteCommand> commandsList = new ArrayList<ExcelWriteCommand>();
                    
                    for (Node commandNode : commandsNodes) {
                        ObjectType cmdObj = (ObjectType) commandNode.evaluate();
                        ExcelWriteCommand excelWriteCommand = new ExcelWriteCommand();
                        
                        for (PairType cmdPair : cmdObj.getPairs()) {
                            String cmdPairKey = cmdPair.getKey().toLowerCase();
                            //System.out.println("cmd for :: " + cmdPairKey);
                            switch (cmdPairKey) {
                                case "\"row\"":
                                    NumberType rowNum = (NumberType) cmdPair.getValue().evaluate();
                                    int row = new Double(rowNum.getDoubleValue()).intValue();
                                    excelWriteCommand.row(row);
                                    break;
                                case "\"cell\"":
                                    int cell = (int) ((NumberType) cmdPair.getValue().evaluate()).getDoubleValue();
                                    excelWriteCommand.cell(cell);
                                    break;
                                case "\"value\"":
                                    OperonValue opValue  = cmdPair.getValue().evaluate();
                                    excelWriteCommand.cellValue(opValue);
                                    break;
                                case "\"sheet\"":
                                    String sheet  = ((StringType) cmdPair.getValue().evaluate()).getJavaStringValue();
                                    excelWriteCommand.sheet(sheet);
                                    break;
                                
                                //
                                // Cell-styling:
                                //
                                
                                case "\"style\"":
                                    String stylename  = ((StringType) cmdPair.getValue().evaluate()).getJavaStringValue();
                                    excelWriteCommand.styleName(stylename);
                                    break;
                                case "\"fill\"":
                                    short fill = (short) ((NumberType) cmdPair.getValue().evaluate()).getDoubleValue();
                                    excelWriteCommand.fill(fill);
                                    break;
                                case "\"bold\"":
                                    Node cs_bold_Node = cmdPair.getValue().evaluate();
                                    if (cs_bold_Node instanceof TrueType) {
                                        excelWriteCommand.bold(true);
                                    }
                                    else {
                                        excelWriteCommand.bold(false);
                                    }
                                    break;
                                case "\"wrap\"":
                                    Node cs_wrap_Node = cmdPair.getValue().evaluate();
                                    if (cs_wrap_Node instanceof TrueType) {
                                        //System.out.println("set wrap true");
                                        excelWriteCommand.wrap(true);
                                    }
                                    else {
                                        //System.out.println("set wrap false");
                                        excelWriteCommand.wrap(false);
                                    }
                                    break;
                                case "\"borderbottom\"":
                                    Node cs_borderBottom_Node = cmdPair.getValue().evaluate();
                                    if (cs_borderBottom_Node instanceof TrueType) {
                                        excelWriteCommand.borderBottom(true);
                                    }
                                    else {
                                        excelWriteCommand.borderBottom(false);
                                    }
                                    break;
                                case "\"bordertop\"":
                                    Node cs_borderTop_Node = cmdPair.getValue().evaluate();
                                    if (cs_borderTop_Node instanceof TrueType) {
                                        excelWriteCommand.borderTop(true);
                                    }
                                    else {
                                        excelWriteCommand.borderTop(false);
                                    }
                                    break;
                                case "\"borderleft\"":
                                    Node cs_borderLeft_Node = cmdPair.getValue().evaluate();
                                    if (cs_borderLeft_Node instanceof TrueType) {
                                        excelWriteCommand.borderLeft(true);
                                    }
                                    else {
                                        excelWriteCommand.borderLeft(false);
                                    }
                                    break;
                                case "\"borderright\"":
                                    Node cs_borderRight_Node = cmdPair.getValue().evaluate();
                                    if (cs_borderRight_Node instanceof TrueType) {
                                        excelWriteCommand.borderRight(true);
                                    }
                                    else {
                                        excelWriteCommand.borderRight(false);
                                    }
                                    break;
                            }
                        }
                        commandsList.add(excelWriteCommand);
                    }
                    info.commands = commandsList;
                    break;
                case "\"load\"":
                    Node load_Node = pair.getValue().evaluate();
                    if (load_Node instanceof TrueType) {
                        info.load = Boolean.TRUE;
                    }
                    else {
                        info.load = Boolean.FALSE;
                    }
                    break;
                case "\"columnwidth\"":
                    int cw = (int) ((NumberType) pair.getValue().evaluate()).getDoubleValue();
                    info.columnWidth = cw;
                    break;
                case "\"landscape\"":
                    Node landscape_Node = pair.getValue().evaluate();
                    if (landscape_Node instanceof TrueType) {
                        info.landscape = Boolean.TRUE;
                    }
                    else {
                        info.landscape = Boolean.FALSE;
                    }
                    break;
                case "\"imagename\"":
                    String imagename = ((StringType) pair.getValue().evaluate()).getJavaStringValue();
                    info.imageName = imagename;
                    break;
                case "\"imageformat\"":
                    String imageformat = ((StringType) pair.getValue().evaluate()).getJavaStringValue();
                    info.imageFormat = imageformat;
                    break;
                case "\"imagexpos\"":
                    int image_xpos = (int) ((NumberType) pair.getValue().evaluate()).getDoubleValue();
                    info.imageXPos = image_xpos;
                    break;
                case "\"imageypos\"":
                    int image_ypos = (int) ((NumberType) pair.getValue().evaluate()).getDoubleValue();
                    info.imageYPos = image_ypos;
                    break;
                case "\"imagescale\"":
                    double image_scale = (double) ((NumberType) pair.getValue().evaluate()).getDoubleValue();
                    info.imageScale = image_scale;
                    break;
                case "\"debug\"":
                    Node debug_Node = pair.getValue().evaluate();
                    if (debug_Node instanceof TrueType) {
                        info.debug = Boolean.TRUE;
                    }
                    else {
                        info.debug = Boolean.FALSE;
                    }
                    break;
                default:
                    log.debug("WriteExcelIntegrationComponent -producer: no mapping for configuration key: " + key);
                    System.err.println("WriteExcelIntegrationComponent -producer: no mapping for configuration key: " + key);
            }
        }
        currentValue.getStatement().setCurrentValue(currentValueCopy);
        return info;
    }
    
    private class Info {
        // File to output
        private String fileName;
        
        private List<ExcelWriteCommand> commands = new ArrayList<ExcelWriteCommand>();
        
        //private List<CellStyle> cellStyles = new ArrayList<CellStyle>();
        
        // false : if file already exists, then open it for augmentation
        // true : if file already exists, then trash it and create new
        private boolean load = Boolean.FALSE;
        
        private Integer columnWidth; // default column-width
        private Boolean landscape = Boolean.FALSE;
        
        //
        // TODO: now supports only one image, should make this an array of objects
        //
        private String imageName = null;
        private String imageFormat = "png";
        private Integer imageXPos = 0;
        private Integer imageYPos = 0;
        private Double imageScale = null;
        
        private Boolean debug = Boolean.FALSE;
    }

    private class ExcelWriteCommand {
        private int row;
        private int cell;
        private OperonValue cellValue;
        private String sheet;
        
        //
        // Cell-styling:
        // TODO: optimize this by gathering the distinct styles
        //       and only creating those style-objects, i.e..
        //       do not create style-object for each cell!
        // This could therefore be a reference to separate style-list.
        // --> that would be more difficult to use! Better to allow per cell-styling
        //     AND reference to cell-style.
        //
        private String styleName; // Each style must be named for internal re-use!
        private Short fill; // cell background color
        private boolean bold = false; // font-bolded
        private boolean borderBottom = false;
        private boolean borderTop = false;
        private boolean borderLeft = false;
        private boolean borderRight = false;
        private boolean wrap = true; // wrap-text on cell
        
        // getters
        public int row() {return row;}
        public int cell() {return cell;}
        public OperonValue cellValue() {return cellValue;}
        public String sheet() {return sheet;}
        
        public String styleName() {return styleName;}
        public Short fill() {return fill;}
        public boolean bold() {return bold;}
        public boolean borderBottom() {return borderBottom;}
        public boolean borderTop() {return borderTop;}
        public boolean borderLeft() {return borderLeft;}
        public boolean borderRight() {return borderRight;}
        public boolean wrap() {return wrap;}
        
        // setters
        public ExcelWriteCommand row(int r) {this.row = r; return this;}
        public ExcelWriteCommand cell(int c) {this.cell = c; return this;}
        public ExcelWriteCommand cellValue(OperonValue v) {this.cellValue = v; return this;}
        
        public ExcelWriteCommand styleName(String sn) {this.styleName = sn; return this;}
        public ExcelWriteCommand sheet(String s) {this.sheet = s; return this;}
        public ExcelWriteCommand fill(short f) {this.fill = f; return this;}
        public ExcelWriteCommand bold(boolean b) {this.bold = b; return this;}
        public ExcelWriteCommand borderBottom(boolean b) {this.borderBottom = b; return this;}
        public ExcelWriteCommand borderTop(boolean b) {this.borderTop = b; return this;}
        public ExcelWriteCommand borderLeft(boolean b) {this.borderLeft = b; return this;}
        public ExcelWriteCommand borderRight(boolean b) {this.borderRight = b; return this;}
        public ExcelWriteCommand wrap(boolean w) {this.wrap = w; return this;}
        
        @Override
        public String toString() {
            return "style=" + this.styleName() + ",value=" + this.cellValue() + ",wrap=" + this.wrap();
        }
    }

}