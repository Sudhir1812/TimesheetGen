package com.example.timesheet.service;

import com.example.timesheet.dto.AttendanceRequest;
import com.example.timesheet.dto.TimesheetRequest;
import com.example.timesheet.entity.Employee;
import com.example.timesheet.repository.EmployeeRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.time.*;
import java.time.format.TextStyle;
import java.util.*;
import java.util.stream.Collectors;

@Service
public class AttendanceService {

    // fixed layout rows
    private static final int HEADER_ROW = 1;
    private static final int EID_ROW = 2;
    private static final int NAME_ROW = 3;
    private static final int SUMMARY_START_ROW = 4; // Days Worked ...
    private static final int REMARKS_ROW = 10;
    private static final int DAILY_START_ROW = 11; // day 1 -> row 11, day 31 -> row 41
    // columns per month block
    private static final int COLS_PER_MONTH = 3; // start, start+1 = label (merged), start+2 = value
    private final EmployeeRepository employeeRepository;

    String empName;
    File file;



    public AttendanceService(EmployeeRepository employeeRepository) {
        this.employeeRepository = employeeRepository;
    }

    /**
     * Writes or updates the month block for given employee in the workbook attendance-<year>.xlsx
     * Returns workbook bytes for immediate download.
     *
     * @return
     */
    public byte[] generateAttendance(AttendanceRequest req) throws Exception {
        // find employee (assuming repository has findByEmployeeId returning Optional<Employee>)
        Employee emp = (Employee) employeeRepository.findByEmployeeId(req.employeeId)
                .orElseThrow(() -> new IllegalArgumentException("Employee not found: " + req.employeeId));

        // parse provided date lists into sets
        Set<LocalDate> leaveSet = toDateSet(req.leaveDates);
        Set<LocalDate> phSet = toDateSet(req.publicHolidays);
        Set<LocalDate> weekOffSet = toDateSet(req.weekOffDates);
        Set<LocalDate> compOffSet = toDateSet(req.compOff);

        YearMonth ym = YearMonth.of(req.year, req.month);
        LocalDate start = ym.atDay(1);
        LocalDate end = ym.atEndOfMonth();

        // workbook filename per year
        String fileName = String.format("attendance-%d.xlsx", req.year);
        file = new File(fileName);

        XSSFWorkbook wb;

        if (file.exists()) {
            try (FileInputStream fis = new FileInputStream(file)) {
                wb = new XSSFWorkbook(fis);
            }
        } else {
            wb = new XSSFWorkbook();
        }

        // sheet name = employee first name (or full) sanitized
         empName = (emp.getEmployeeName() == null ? "Employee" : emp.getEmployeeName().split(" ")[0]);
        String sheetName = sanitizeSheetName(empName);
        if (sheetName.length() > 31) sheetName = sheetName.substring(0, 31);

        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) {
            sheet = wb.createSheet(sheetName);
            // set default row heights to make space


            for (int i = 1; i < DAILY_START_ROW + ym.lengthOfMonth(); i++) {
                Row rr = sheet.getRow(i);
                if (rr == null) rr = sheet.createRow(i);
                //rr.setHeightInPoints(18f);
            }
        }

        // create styles (fresh per invocation is fine)
        Map<String, CellStyle> styles = createStyles(wb);

        // compute start column for month (0-based)
        int monthIndex = req.month - 1;
        if (monthIndex < 0 || monthIndex > 11) {
            throw new IllegalArgumentException("month must be 1..12");
        }

        int startCol = (monthIndex * COLS_PER_MONTH) + 1;
        int labelCol2 = startCol + 1;
        int valueCol = startCol + 2;

        // remove any merged regions or old data in this month block (rows HEADER_ROW .. DAILY_START_ROW+30)
        clearMonthBlock(sheet, HEADER_ROW, (DAILY_START_ROW + ym.lengthOfMonth()) - 1, startCol, valueCol);

        // write month header (merged across the 3 columns)
        Row header = getOrCreateRow(sheet, HEADER_ROW);
        Cell monthCell = header.createCell(startCol);
        ym = YearMonth.of(req.year, req.month);
        monthCell.setCellValue(ym.getMonth().getDisplayName(TextStyle.FULL, Locale.ENGLISH) + "-" + String.valueOf(req.year).substring(2));
        monthCell.setCellStyle(styles.get("monthHeader"));
        sheet.addMergedRegion(new CellRangeAddress(HEADER_ROW, HEADER_ROW, startCol, valueCol));
        // ensure style across merged cells
        applyStyleToRegionCells(sheet, HEADER_ROW, HEADER_ROW, startCol, valueCol, styles.get("monthHeader"));

        // EID row (label merged across labelCol1..labelCol2, value in valueCol)
        writeLabelValueInBlock(sheet, EID_ROW, startCol, labelCol2, valueCol, "EID", String.valueOf(emp.getEmployeeId()), styles);

        // Employee name row
        writeLabelValueInBlock(sheet, NAME_ROW, startCol, labelCol2, valueCol, "Employee Name", emp.getEmployeeName(), styles);

        // Prepare daily statuses for this month
        YearMonth yearMonth = YearMonth.of(req.year, req.month);
        int daysInMonth = yearMonth.lengthOfMonth();


        // compute daily statuses and counts
        List<DayStatus> dayStatusList = new ArrayList<>();

        int sumP = 0, sumL = 0, sumPH = 0, sumWO = 0, sumCO = 0;
        for (LocalDate d = start; !d.isAfter(end); d = d.plusDays(1)) {
            boolean isSunday = d.getDayOfWeek() == DayOfWeek.SUNDAY;
            boolean isSaturday = d.getDayOfWeek() == DayOfWeek.SATURDAY;
            int saturdayNumber = 0;
            int dom = d.getDayOfMonth();
            if (isSaturday) saturdayNumber = ((dom - 1) / 7) + 1;

            if ((req.saturdayWeekoff == TimesheetRequest.WEEKOFF_1ST_3RD || req.saturdayWeekoff == TimesheetRequest.WEEKOFF_2ND_4TH || req.saturdayWeekoff == TimesheetRequest.ALL_WEEKOFF) && isSaturday) {
                //int saturdayNumber = ((dom - 1) / 7) + 1;
                if ((req.saturdayWeekoff == TimesheetRequest.WEEKOFF_1ST_3RD && (saturdayNumber == 1 || saturdayNumber == 3)) ||
                        (req.saturdayWeekoff == TimesheetRequest.WEEKOFF_2ND_4TH && (saturdayNumber == 2 || saturdayNumber == 4)) ||
                        (req.saturdayWeekoff == TimesheetRequest.ALL_WEEKOFF)) {

                    if (phSet.contains(d) || leaveSet.contains(d) || compOffSet.contains(d)) {
                        throw new IllegalArgumentException("Saturday " + d + " cannot be marked as leave or holidays or compOff");
                    }
                }

            }

            String status;
            if (phSet.contains(d)) {
                status = "PH";
                sumPH++;
            } else if (leaveSet.contains(d)) {
                status = "L";
                sumL++;
            } else if (isSunday) {
                if (weekOffSet.contains(d)) {
                    status = "P";
                    sumP++;
                    sumWO++;
                } else {
                    status = "WO";
                }
            } else if (req.saturdayWeekoff == AttendanceRequest.WEEKOFF_1ST_3RD && isSaturday && (saturdayNumber == 1 || saturdayNumber == 3)) {
                if (weekOffSet.contains(d)) {
                    status = "P";
                    sumP++;
                    sumWO++;
                } else {
                    status = "WO";
                }
            } else if (req.saturdayWeekoff == AttendanceRequest.WEEKOFF_2ND_4TH && isSaturday && (saturdayNumber == 2 || saturdayNumber == 4)) {
                if (weekOffSet.contains(d)) {
                    status = "P";
                    sumP++;
                    sumWO++;
                } else {
                    status = "WO";
                }
            } else if (req.saturdayWeekoff == AttendanceRequest.ALL_WEEKOFF && isSaturday) {
                if (weekOffSet.contains(d)) {
                    status = "P";
                    sumP++;
                    sumWO++;
                } else {
                    status = "WO";
                }
            } else if (compOffSet.contains(d)) {
                status = "CO";
                sumCO++;
            } else {
                status = "P";
                sumP++;
            }
            dayStatusList.add(new DayStatus(d, status, d.getDayOfWeek()));
        }


        // Write summary rows
        writeLabelValueInBlock(sheet, SUMMARY_START_ROW, startCol, labelCol2, valueCol, "Days Worked", String.valueOf(sumP), styles);
        writeLabelValueInBlock(sheet, SUMMARY_START_ROW + 1, startCol, labelCol2, valueCol, "Total Working Days", String.valueOf((sumP + sumL + sumCO)-sumWO), styles);
        writeLabelValueInBlock(sheet, SUMMARY_START_ROW + 2, startCol, labelCol2, valueCol, "Leaves", String.valueOf(sumL), styles);
        writeLabelValueInBlock(sheet, SUMMARY_START_ROW + 3, startCol, labelCol2, valueCol, "WO Worked", String.valueOf(sumWO), styles);
        writeLabelValueInBlock(sheet, SUMMARY_START_ROW + 4, startCol, labelCol2, valueCol, "Comp Off", String.valueOf(sumCO), styles);
        writeLabelValueInBlock(sheet, SUMMARY_START_ROW + 5, startCol, labelCol2, valueCol, "PH", String.valueOf(sumPH), styles);

        // Remarks
        writeLabelValueInBlock(sheet, REMARKS_ROW, startCol, labelCol2, valueCol, "Remarks", req.remarks == null ? "" : req.remarks, styles);

        // Daily rows: use DAILY_START_ROW  DAILY_START_ROW + 30 (1..31)
        for (int i = 0; i < ym.lengthOfMonth(); i++) {
            int writeRow = DAILY_START_ROW + i;
            Row dr = getOrCreateRow(sheet, writeRow);

            // label merged cell shows date + day (for days within the month), otherwise blank
            Cell lbl = dr.createCell(startCol);
            lbl.setCellStyle(styles.get("dayCell"));
            //sheet.addMergedRegion(new CellRangeAddress(writeRow, writeRow, startCol, labelCol2));
            // ensure partner cell has style
            Cell partner = dr.getCell(labelCol2);
            if (partner == null) partner = dr.createCell(labelCol2);
            partner.setCellStyle(styles.get("dayCell"));

            Cell val = dr.createCell(valueCol);
            val.setCellStyle(styles.get("statusCell"));

            if (i < daysInMonth) {
                DayStatus ds = dayStatusList.get(i);
                String dateLabel = String.format("%02d-%s", ds.date.getDayOfMonth(),
                        ds.date.getMonth().getDisplayName(TextStyle.SHORT_STANDALONE, Locale.ENGLISH));
                lbl.setCellValue(dateLabel);
                // put date and day on separate lines (Excel will show newline if wrap is enabled)
                partner.setCellValue(ds.dayOfWeek.getDisplayName(TextStyle.SHORT_STANDALONE, Locale.ENGLISH));
                val.setCellValue(ds.status);
            } else {
                // empty for days beyond month length
                lbl.setCellValue("");
                val.setCellValue("");
            }
        }

        // Apply outer border for this month block (optional: double border)
        int bottomRow = DAILY_START_ROW + ym.lengthOfMonth() - 1;
        CellRangeAddress outer = new CellRangeAddress(HEADER_ROW, bottomRow, startCol, valueCol);
        RegionUtil.setBorderTop(BorderStyle.THICK, outer, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THICK, outer, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THICK, outer, sheet);
        RegionUtil.setBorderRight(BorderStyle.THICK, outer, sheet);
        RegionUtil.setTopBorderColor(IndexedColors.BLACK.getIndex(), outer, sheet);
        RegionUtil.setBottomBorderColor(IndexedColors.BLACK.getIndex(), outer, sheet);
        RegionUtil.setLeftBorderColor(IndexedColors.BLACK.getIndex(), outer, sheet);
        RegionUtil.setRightBorderColor(IndexedColors.BLACK.getIndex(), outer, sheet);

        // auto-size a few columns for the whole sheet (only up to used columns)
        int usedCols = COLS_PER_MONTH * 12;
        for (int c = 0; c < usedCols; c++) {
            sheet.autoSizeColumn(c);
        }


        // reorder methods for better structuring
        // (e.g. utility methods at the end)
        // (also, add blank lines between groups of methods)

        // buildTimesheetXLSX
        // getOrCreateRow
        // writeLabelValueInBlock
        // createStyles
        // applyStyleToRegionCells
        // save workbook back to disk
        try (FileOutputStream fos = new FileOutputStream(file)) {
            wb.write(fos);
        }

          // MUST CLOSE

// ---- Return workbook bytes if needed ----
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (baos) {
            wb.write(baos);
        }
        wb.close();
        return baos.toByteArray();
    }

    // ---------------------- helper utilities ----------------------

    private Row getOrCreateRow(Sheet sheet, int r) {
        Row row = sheet.getRow(r);
        if (row == null) row = sheet.createRow(r);
        return row;
    }

    private void writeLabelValueInBlock(Sheet sheet,
                                        int rowIndex,
                                        int labelCol1,
                                        int labelCol2,
                                        int valueCol,
                                        String label,
                                        String value,
                                        Map<String, CellStyle> styles) {
        Row row = getOrCreateRow(sheet, rowIndex);

        // label (merged labelCol1..labelCol2)
        Cell labelCell = row.getCell(labelCol1);
        if (labelCell == null) labelCell = row.createCell(labelCol1);
        labelCell.setCellValue(label);
        labelCell.setCellStyle(styles.get("label"));

        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, labelCol1, labelCol2));

        // ensure partner has style too
        Cell partner = row.getCell(labelCol2);
        if (partner == null) partner = row.createCell(labelCol2);
        partner.setCellStyle(styles.get("label"));

        // value (single column)
        Cell v = row.getCell(valueCol);
        if (v == null) v = row.createCell(valueCol);
        v.setCellValue(value);
        v.setCellStyle(styles.get("value"));
    }

    // remove merged regions and clear cell values in the rectangle block
    private void clearMonthBlock(Sheet sheet, int rowStart, int rowEnd, int colStart, int colEnd) {
        // remove merged regions that intersect block
        List<CellRangeAddress> toRemove = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress cra = sheet.getMergedRegion(i);
            if (!(cra.getLastRow() < rowStart || cra.getFirstRow() > rowEnd
                    || cra.getLastColumn() < colStart || cra.getFirstColumn() > colEnd)) {
                toRemove.add(cra);
            }
        }
        // physically remove by index (work backwards)
        for (CellRangeAddress cra : toRemove) {
            // find index
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                if (sheet.getMergedRegion(i).formatAsString().equals(cra.formatAsString())) {
                    sheet.removeMergedRegion(i);
                    break;
                }
            }
        }

        // clear content & styles in region
        for (int r = rowStart; r <= rowEnd; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = colStart; c <= colEnd; c++) {
                Cell cell = row.getCell(c);
                if (cell != null) {
                    cell.setCellStyle(null);
                    cell.setCellValue("");
                }
            }
        }
    }

    private void applyStyleToRegionCells(Sheet sheet, int rowStart, int rowEnd, int colStart, int colEnd, CellStyle style) {
        for (int r = rowStart; r <= rowEnd; r++) {
            Row row = getOrCreateRow(sheet, r);
            for (int c = colStart; c <= colEnd; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) cell = row.createCell(c);
                cell.setCellStyle(style);
            }
        }
    }

    private Set<LocalDate> toDateSet(List<String> list) {
        if (list == null) return Collections.emptySet();
        return list.stream().map(LocalDate::parse).collect(Collectors.toCollection(LinkedHashSet::new));
    }

    private String sanitizeSheetName(String name) {
        if (name == null) return "Sheet";
        return name.replaceAll("[\\\\/?*\\[\\]]", "_");
    }

    // create commonly used styles
    private Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> s = new HashMap<>();

        // monthHeader
        Font mh = wb.createFont();
        mh.setBold(true);
        mh.setFontHeightInPoints((short) 12);
        CellStyle monthHeader = wb.createCellStyle();
        monthHeader.setFont(mh);
        monthHeader.setAlignment(HorizontalAlignment.CENTER);
        monthHeader.setVerticalAlignment(VerticalAlignment.CENTER);
        monthHeader.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        monthHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        monthHeader.setBorderTop(BorderStyle.THICK);
        monthHeader.setBorderBottom(BorderStyle.THICK);
        monthHeader.setBorderLeft(BorderStyle.THICK);
        monthHeader.setBorderRight(BorderStyle.THICK);
        s.put("monthHeader", monthHeader);

        // label style (left two columns)
        Font lf = wb.createFont();
        lf.setBold(true);
        CellStyle label = wb.createCellStyle();
        label.setFont(lf);
        label.setAlignment(HorizontalAlignment.CENTER);
        label.setVerticalAlignment(VerticalAlignment.CENTER);
        //label.setWrapText(true);
        label.setBorderTop(BorderStyle.THIN);
        label.setBorderBottom(BorderStyle.THIN);
        label.setBorderLeft(BorderStyle.THIN);
        label.setBorderRight(BorderStyle.THIN);
        s.put("label", label);

        // value style (right column)
        CellStyle value = wb.createCellStyle();
        Font valueFont = wb.createFont();
        valueFont.setBold(true);
        value.setFont(valueFont);
        value.setAlignment(HorizontalAlignment.CENTER);
        value.setVerticalAlignment(VerticalAlignment.CENTER);
        //value.setWrapText(true);
        value.setBorderTop(BorderStyle.THIN);
        value.setBorderBottom(BorderStyle.THIN);
        value.setBorderLeft(BorderStyle.THIN);
        value.setBorderRight(BorderStyle.THIN);
        s.put("value", value);

        // dayCell (label for daily rows)
        CellStyle dayCell = wb.createCellStyle();
        Font dayFont = wb.createFont();
        dayFont.setBold(true);
        dayCell.setFont(dayFont);
        dayCell.setAlignment(HorizontalAlignment.CENTER);
        dayCell.setVerticalAlignment(VerticalAlignment.CENTER);
        //dayCell.setWrapText(true);
        dayCell.setBorderTop(BorderStyle.THIN);
        dayCell.setBorderBottom(BorderStyle.THIN);
        dayCell.setBorderLeft(BorderStyle.THIN);
        dayCell.setBorderRight(BorderStyle.THIN);
        s.put("dayCell", dayCell);

        // status cell for daily value
        CellStyle statusCell = wb.createCellStyle();
        statusCell.setAlignment(HorizontalAlignment.CENTER);
        statusCell.setVerticalAlignment(VerticalAlignment.CENTER);
        statusCell.setBorderTop(BorderStyle.THIN);
        statusCell.setBorderBottom(BorderStyle.THIN);
        statusCell.setBorderLeft(BorderStyle.THIN);
        statusCell.setBorderRight(BorderStyle.THIN);
        s.put("statusCell", statusCell);

        return s;
    }

    public byte[] extractSheetsForEmployee() throws Exception {



        XSSFWorkbook sourceWorkbook;

        try (FileInputStream fis = new FileInputStream(file)) {
                sourceWorkbook = new XSSFWorkbook(fis);
            }

        // Create new workbook in memory
        XSSFWorkbook destWorkbook = new XSSFWorkbook();

        for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); i++) {

            Sheet src = sourceWorkbook.getSheetAt(i);
            String name = src.getSheetName();

            // match Sudhir, Sudip, etc.
            if (name.equals(empName)) {
                Sheet dest = destWorkbook.createSheet(name);
                copySheetSafe(src, dest, destWorkbook);

                // auto-size a few columns for the whole sheet (only up to used columns)
                int usedCols = COLS_PER_MONTH * 12;
                for (int c = 0; c < usedCols; c++) {
                    dest.autoSizeColumn(c);
                }
            }
        }

        // return bytes
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            destWorkbook.write(bos);
            destWorkbook.close();
            sourceWorkbook.close();
            return bos.toByteArray();
        }
    }


    private void copySheetSafe(Sheet src, Sheet dest, XSSFWorkbook newWb) {

        // 1️⃣ Column widths
        Row firstRow = src.getRow(0);
        if (firstRow != null) {
            for (int c = 0; c < firstRow.getLastCellNum(); c++) {
                dest.setColumnWidth(c, src.getColumnWidth(c));
            }
        }

        // 2️⃣ Merged regions
        for (int i = 0; i < src.getNumMergedRegions(); i++) {
            dest.addMergedRegion(src.getMergedRegion(i));
        }

        // 3️⃣ Rows + cells
        for (int r = 0; r <= src.getLastRowNum(); r++) {

            Row srcRow = src.getRow(r);
            if (srcRow == null) continue;

            Row destRow = dest.createRow(r);
            destRow.setHeight(srcRow.getHeight());

            for (int c = 0; c < srcRow.getLastCellNum(); c++) {

                Cell oldCell = srcRow.getCell(c);
                if (oldCell == null) continue;

                Cell newCell = destRow.createCell(c);

                // Copy FULL style including fonts
                newCell.setCellStyle(cloneCellStyle(newWb, oldCell.getCellStyle()));

                // Comments
                if (oldCell.getCellComment() != null)
                    newCell.setCellComment(oldCell.getCellComment());

                // Hyperlinks
                if (oldCell.getHyperlink() != null)
                    newCell.setHyperlink(oldCell.getHyperlink());

                // Value
                switch (oldCell.getCellType()) {
                    case STRING -> newCell.setCellValue(oldCell.getRichStringCellValue());
                    case NUMERIC -> newCell.setCellValue(oldCell.getNumericCellValue());
                    case BOOLEAN -> newCell.setCellValue(oldCell.getBooleanCellValue());
                    case FORMULA -> newCell.setCellFormula(oldCell.getCellFormula());
                    case BLANK -> newCell.setBlank();
                    default -> { }
                }
            }
        }
    }



    private CellStyle cloneCellStyle(XSSFWorkbook wb, CellStyle oldStyle) {

        if (oldStyle == null) {
            return wb.createCellStyle();
        }

        XSSFCellStyle newStyle = wb.createCellStyle();
        newStyle.cloneStyleFrom(oldStyle); // copies: borders, fills, alignments, etc.

        // Copy font
        XSSFCellStyle oldXssf = (XSSFCellStyle) oldStyle;
        XSSFFont oldFont = oldXssf.getFont();

        XSSFFont newFont = wb.createFont();
        newFont.setBold(oldFont.getBold());
        newFont.setColor(oldFont.getColor());
        newFont.setFontHeight(oldFont.getFontHeight());
        newFont.setFontName(oldFont.getFontName());
        newFont.setItalic(oldFont.getItalic());
        newFont.setUnderline(oldFont.getUnderline());
        newFont.setStrikeout(oldFont.getStrikeout());
        newFont.setTypeOffset(oldFont.getTypeOffset());

        newStyle.setFont(newFont);

        return newStyle;
    }




    // DayStatus
    private static class DayStatus {
        LocalDate date;
        String status;
        DayOfWeek dayOfWeek;

        DayStatus(LocalDate date, String status, DayOfWeek dow) {
            this.date = date;
            this.status = status;
            this.dayOfWeek = dow;
        }
    }


}
