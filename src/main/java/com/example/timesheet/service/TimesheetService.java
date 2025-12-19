package com.example.timesheet.service;

import com.example.timesheet.dto.TimesheetRequest;
import com.example.timesheet.dto.TimesheetRow;
import com.example.timesheet.entity.Employee;
import com.example.timesheet.repository.EmployeeRepository;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.*;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.*;
import java.util.function.Consumer;

@Service
public class TimesheetService {

    private final EmployeeRepository employeeRepository;
    private final DateTimeFormatter iso = DateTimeFormatter.ISO_LOCAL_DATE;
    private static final String IN_TIME = "08:00";
    private static final String OUT_TIME = "17:00";

    public TimesheetService(EmployeeRepository employeeRepository) {
        this.employeeRepository = employeeRepository;
    }

    File excelFile;
    String sheetName;

    public byte[] generateTimesheet(TimesheetRequest req) throws Exception {


        Employee emp = (Employee) employeeRepository.findByEmployeeId(req.getEmployeeId())
                .orElseThrow(() -> new IllegalArgumentException("Employee not found: " + req.getEmployeeId()));

        Set<LocalDate> holidaySet = new TreeSet<>();
        Set<LocalDate> leaveSet = new TreeSet<>();


        // validate holidays and leave dates belong to the same month and year as in the request
        if (req.getHolidays() != null) {
            for (String d : req.getHolidays()) {
                LocalDate holiday = LocalDate.parse(d, iso);
                if (holiday.getYear() != req.getYear() || holiday.getMonthValue() != req.getMonth()) {
                    throw new IllegalArgumentException("Holiday " + d + " doesn't belong to month " + req.getMonth() + "/" + req.getYear());
                }
                holidaySet.add(holiday);
            }
        }

        if (req.getLeaveDates() != null) {
            for (String d : req.getLeaveDates()) {
                LocalDate leave = LocalDate.parse(d, iso);
                if (leave.getYear() != req.getYear() || leave.getMonthValue() != req.getMonth()) {
                    throw new IllegalArgumentException("Leave " + d + " doesn't belong to month " + req.getMonth() + "/" + req.getYear());
                }
                leaveSet.add(leave);
            }
        }


        YearMonth ym = YearMonth.of(req.getYear(), req.getMonth());
        LocalDate start = ym.atDay(1);
        LocalDate end = ym.atEndOfMonth();

        List<TimesheetRow> rows = new ArrayList<>();
        int regularWorkDays = 0;
        int leaveDays = 0;
        double totalHours = 0.0;

        for (LocalDate date = start; !date.isAfter(end); date = date.plusDays(1)) {
            TimesheetRow row = new TimesheetRow();
            row.dayName = date.getDayOfWeek().getDisplayName(TextStyle.SHORT, Locale.ENGLISH);
            row.date = String.format("%02d", date.getDayOfMonth());
            row.managerApproval = req.getManagerApproval();
            row.remarks = holidaySet.contains(date)?req.getRemarks().get(new ArrayList<>(holidaySet).indexOf(date)): "";

            boolean isSunday = date.getDayOfWeek() == DayOfWeek.SUNDAY;
            boolean isSaturday = date.getDayOfWeek() == DayOfWeek.SATURDAY;
            int saturdayNumber = 0;
            int dom = date.getDayOfMonth();
            if (isSaturday) saturdayNumber = ((dom - 1) / 7) + 1;


            if ((req.getSaturdayWeekoff() == TimesheetRequest.WEEKOFF_1ST_3RD || req.getSaturdayWeekoff() == TimesheetRequest.WEEKOFF_2ND_4TH || req.getSaturdayWeekoff() == TimesheetRequest.ALL_WEEKOFF) && isSaturday) {
                //int saturdayNumber = ((dom - 1) / 7) + 1;
                if ((req.getSaturdayWeekoff() == TimesheetRequest.WEEKOFF_1ST_3RD && (saturdayNumber == 1 || saturdayNumber == 3)) ||
                        (req.getSaturdayWeekoff() == TimesheetRequest.WEEKOFF_2ND_4TH && (saturdayNumber == 2 || saturdayNumber == 4)) ||
                        (req.getSaturdayWeekoff() == TimesheetRequest.ALL_WEEKOFF)) {

                    if (holidaySet.contains(date) || leaveSet.contains(date)) {
                        throw new IllegalArgumentException("Saturday " + date + " cannot be marked as leave or holidays");
                    }
                }


            }

            if (leaveSet.contains(date)) {
                row.activity = "Leave";
                row.inTime = "";
                row.outTime = "";
                row.duration = "";
                leaveDays++;
            } else if (isSunday) {
                row.activity = "WeekOff";
                row.inTime = "";
                row.outTime = "";
                row.duration = "";
            } else if (holidaySet.contains(date)) {
                row.activity = "Holiday";
                row.inTime = "";
                row.outTime = "";
                row.duration = "";
            }
            else if (req.getSaturdayWeekoff() == TimesheetRequest.WEEKOFF_1ST_3RD && isSaturday && (saturdayNumber == 1 || saturdayNumber == 3)) {
                row.activity = "WeekOff";
                row.inTime = "";
                row.outTime = "";
                row.duration = "";

            } else if (req.getSaturdayWeekoff() == TimesheetRequest.WEEKOFF_2ND_4TH && isSaturday && (saturdayNumber == 2 || saturdayNumber == 4)) {
                row.activity = "WeekOff";
                row.inTime = "";
                row.outTime = "";
                row.duration = "";
            } else if (req.getSaturdayWeekoff() == TimesheetRequest.ALL_WEEKOFF && isSaturday ) {
            row.activity = "WeekOff";
            row.inTime = "";
            row.outTime = "";
            row.duration = "";
            } else {
                row.activity = "Regular Work";
                row.inTime = IN_TIME;
                double hrs;
                if (isSaturday) {
                    row.outTime = "13:00";
                    hrs = Duration.between(LocalTime.parse(IN_TIME), LocalTime.parse("13:00")).toMinutes() / 60.0;
                } else {
                    row.outTime = OUT_TIME;
                    hrs = Duration.between(LocalTime.parse(IN_TIME), LocalTime.parse(OUT_TIME)).toMinutes() / 60.0;
                }
                row.duration = String.format("%.2f", hrs);
                totalHours += hrs;
                regularWorkDays++;
            }
            rows.add(row);
        }

        return buildExcel(emp, req.getYear(), req.getMonth(), rows, regularWorkDays, leaveDays, totalHours);
    }

    private byte[] buildExcel(Employee emp, int year, int month,
                              List<TimesheetRow> rows, int regularWorkDays, int leaveDays, double totalHours) throws Exception {


//        String fileName = String.format("timesheet-%d-%02d.xlsx", year, month);
        String fileName = String.format("timesheet-%d.xlsx", year);
         excelFile = new File(fileName);

        XSSFWorkbook wb;

        // If file exists open it, else create new
        if (excelFile.exists()) {
            try (FileInputStream fis = new FileInputStream(excelFile)) {
                wb = new XSSFWorkbook(fis);
            }
        } else {
            wb = new XSSFWorkbook();
        }


        Month monthEnum = Month.of(month);
        String monthName = monthEnum.getDisplayName(TextStyle.SHORT_STANDALONE, Locale.ENGLISH);
        // Sheet name e.g.: Sudhir_2025_01
         sheetName = emp.getEmployeeName().split(" ")[0] + ", " + monthName + " " + year;
        int existingIndex = wb.getSheetIndex(sheetName);
        if (existingIndex != -1) {
            wb.removeSheetAt(existingIndex);
        }

        Sheet sheet = wb.createSheet(sheetName);



            // ---- Styles ----

            CellStyle borderDashDot = wb.createCellStyle();
            borderDashDot.setBorderTop(BorderStyle.DASHED);
            borderDashDot.setBorderBottom(BorderStyle.DASHED);
            borderDashDot.setBorderLeft(BorderStyle.DASHED);
            borderDashDot.setBorderRight(BorderStyle.DASHED);

            // Set border color
            short lightGrayDashed = IndexedColors.GREY_25_PERCENT.getIndex();
            borderDashDot.setTopBorderColor(lightGrayDashed);
            borderDashDot.setBottomBorderColor(lightGrayDashed);
            borderDashDot.setLeftBorderColor(lightGrayDashed);
            borderDashDot.setRightBorderColor(lightGrayDashed);


            CellStyle doubleBorder = wb.createCellStyle();
            doubleBorder.setBorderTop(BorderStyle.DOUBLE);
            doubleBorder.setBorderBottom(BorderStyle.DOUBLE);
            doubleBorder.setBorderLeft(BorderStyle.DOUBLE);
            doubleBorder.setBorderRight(BorderStyle.DOUBLE);

            CellStyle doubleBorderGrey = wb.createCellStyle();
            doubleBorderGrey.cloneStyleFrom(doubleBorder);
            doubleBorderGrey.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            doubleBorderGrey.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle lightGray = wb.createCellStyle();
            lightGray.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            lightGray.setFillPattern(FillPatternType.SOLID_FOREGROUND);


            // Title style (double border + underline + bold + centered)
            Font titleFont = wb.createFont();
            titleFont.setBold(true);
            titleFont.setUnderline(FontUnderline.SINGLE.getByteValue());
            titleFont.setFontHeightInPoints((short) 16);

            CellStyle titleStyle = wb.createCellStyle();
            titleStyle.cloneStyleFrom(doubleBorder);
            titleStyle.setFont(titleFont);
            titleStyle.setAlignment(HorizontalAlignment.CENTER);
            titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // Label / value styles for merged rows
            Font labelF = wb.createFont();
            labelF.setBold(true);
            labelF.setItalic(true);
            CellStyle labelStyle = wb.createCellStyle();
            labelStyle.setFont(labelF);
            labelStyle.setAlignment(HorizontalAlignment.RIGHT);


            CellStyle valueStyle = wb.createCellStyle();
            valueStyle.setAlignment(HorizontalAlignment.LEFT);

            // Year/Month header style (bold black + grey + double border)
            Font boldBlack = wb.createFont();
            boldBlack.setBold(true);
            boldBlack.setColor(IndexedColors.BLACK.getIndex());

            CellStyle yearMonthHeaderStyle = wb.createCellStyle();
            yearMonthHeaderStyle.cloneStyleFrom(doubleBorderGrey);
            yearMonthHeaderStyle.setFont(boldBlack);
            yearMonthHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
            yearMonthHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // Year/month value style (double border, centered)
            CellStyle yearMonthValueStyle = wb.createCellStyle();
            yearMonthValueStyle.cloneStyleFrom(doubleBorder);
            yearMonthValueStyle.setAlignment(HorizontalAlignment.CENTER);
            yearMonthValueStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // GRSE style
            Font whiteBold = wb.createFont();
            whiteBold.setBold(true);
            whiteBold.setColor(IndexedColors.WHITE.getIndex());

            CellStyle grseStyle = wb.createCellStyle();
            grseStyle.cloneStyleFrom(doubleBorderGrey);
            grseStyle.setFont(whiteBold);
            grseStyle.setAlignment(HorizontalAlignment.CENTER);
            grseStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            int r = 0;

            // Helper lambdas
            // Helper lambda — explicit type so Java can infer
            Consumer<Row> makeFirstCellBlank = row -> {
                Cell c = row.createCell(0);
                c.setCellValue("        ");
            };


            // ROW 1 blank
            Row row1 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row1);

            // ROW 2 title (B->I merged), double border, underlined
            Row row2 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row2);
            Cell title = row2.createCell(1);
            title.setCellValue("TIMESHEET CALCULATOR");
            title.setCellStyle(titleStyle);
            sheet.addMergedRegion(new CellRangeAddress(row2.getRowNum(), row2.getRowNum(), 1, 8));
            // apply style across merged region
            for (int cc = 2; cc <= 8; cc++) {
                Cell c = row2.getCell(cc);
                if (c == null) c = row2.createCell(cc);
                c.setCellStyle(titleStyle);
            }

            // Insert Company Logo

            // -----------------------------
        InputStream logoStream = getClass().getResourceAsStream("/logo.png");
        if (logoStream == null) {
            throw new RuntimeException("logo.png not found in resources folder");
        }
        byte[] imageBytes = IOUtils.toByteArray(logoStream);
        int imageId = wb.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);


            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFClientAnchor logoAnchor = new XSSFClientAnchor(
                    0, 20000, 0, -20000,
                    3, 1,   // B2
                    4, 2    // small area under control
            );
            XSSFPicture pic = drawing.createPicture(logoAnchor, imageId);


            sheet.getRow(1).setHeightInPoints(30);

            // Rows 3 & 4 blank
            Row row3 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row3);
            Row row4 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row4);

            // Row 5 blank
            Row row5 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row5);

            // EMPLOYEE INFO BLOCK STARTS at current r (this is row index 5 as 0-based)
            int empBlockStart = r;

            // create merged labeled rows (B-C and D-I) keeping cell A blank
            createMergedLabeledRow(sheet, r++, "Employee Name:", emp.getEmployeeName() == null ? "" : emp.getEmployeeName(), wb, labelStyle, valueStyle);
            createMergedLabeledRow(sheet, r++, "Employee ID:", emp.getEmployeeId() == null ? "" : (emp.getEmployeeId()), wb, labelStyle, valueStyle);
            createMergedLabeledRow(sheet, r++, "Email id:", emp.getEmail() == null ? "" : emp.getEmail(), wb, labelStyle, valueStyle);
            createMergedLabeledRow(sheet, r++, "Joining Date:", emp.getJoiningDate() == null ? "" : emp.getJoiningDate(), wb, labelStyle, valueStyle);

            // Row 10 blank
            Row row10 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row10);

            // Stats
            createMergedLabeledRow(sheet, r++, "Regular Work", String.valueOf(regularWorkDays), wb, labelStyle, valueStyle);
            createMergedLabeledRow(sheet, r++, "Total Hours", String.format("%.2f", totalHours), wb, labelStyle, valueStyle);
            createMergedLabeledRow(sheet, r++, "Leave", String.valueOf(leaveDays), wb, labelStyle, valueStyle);

            // Row 14 blank

            Row row14 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row14);

            for (int row = 4; row <= 13; row++) {   // 5th row is index 4
                Row r1 = sheet.getRow(row);
                if (r1 == null) r1 = sheet.createRow(row);

                for (int col = 1; col <= 8; col++) {  // B → I
                    Cell c = r1.getCell(col);
                    if (c == null) c = r1.createCell(col);

                    patchGrayAndDoubleBorder(c, wb);
                }
            }



            CellRangeAddress outerBlock = new CellRangeAddress(4, 13, 1, 8);  // B5 → I14

            RegionUtil.setBorderTop(BorderStyle.DOUBLE, outerBlock, sheet);
            RegionUtil.setBorderBottom(BorderStyle.DOUBLE, outerBlock, sheet);
            RegionUtil.setBorderLeft(BorderStyle.DOUBLE, outerBlock, sheet);
            RegionUtil.setBorderRight(BorderStyle.DOUBLE, outerBlock, sheet);




            // Rows 15 & 16 blank
            Row row15 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row15);
            Row row16 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row16);

            // Row 17: YEAR and MONTH headers (F & G)
            Row row17 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row17);
            Cell yHead = row17.createCell(5);
            yHead.setCellValue("YEAR");
            yHead.setCellStyle(yearMonthHeaderStyle);
            Cell mHead = row17.createCell(6);
            mHead.setCellValue("MONTH");
            mHead.setCellStyle(yearMonthHeaderStyle);

            // Row 18: YEAR & MONTH values (double boxed)
            Row row18 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row18);
            Cell yVal = row18.createCell(5);
            yVal.setCellValue(year);
            yVal.setCellStyle(yearMonthValueStyle);
            Cell mVal = row18.createCell(6);
            mVal.setCellValue(Month.of(month).getDisplayName(TextStyle.FULL, Locale.ENGLISH));
            mVal.setCellStyle(yearMonthValueStyle);

            // Row 19: GRSE in H (index 7)
            Row row19 = sheet.createRow(r++);
            makeFirstCellBlank.accept(row19);
            Cell grseCell = row19.createCell(6);
            grseCell.setCellValue("GRSE");
            grseCell.setCellStyle(grseStyle);
            grseCell.setCellStyle(grseStyle);




            // Table header (next row)
            String[] headers = {"Day", "Date", "In Time", "Out Time", "Duration  (Hrs)", "Activity", "Manager Approval", "Remarks"};
            Row headerRow = sheet.createRow(r++);
            makeFirstCellBlank.accept(headerRow);
            Font hf = wb.createFont();
            hf.setBold(true);
            CellStyle headerStyle = wb.createCellStyle();
            headerStyle.setFont(hf);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            for (int i = 0; i < headers.length; i++) {
                Cell c = headerRow.createCell(i + 1);
                c.setCellValue(headers[i]);
                c.setCellStyle(headerStyle);
                if (i == 1) {
                    sheet.setColumnWidth(i + 1, 1000); // date col is wider
                }
            }

            CellRangeAddress outerBlock2 = new CellRangeAddress(r-1, r-1, 1, 8);  // B5 → I14

            RegionUtil.setBorderTop(BorderStyle.DOUBLE, outerBlock2, sheet);
            RegionUtil.setBorderBottom(BorderStyle.DOUBLE, outerBlock2, sheet);
            RegionUtil.setBorderLeft(BorderStyle.DOUBLE, outerBlock2, sheet);
            RegionUtil.setBorderRight(BorderStyle.DOUBLE, outerBlock2, sheet);




            // Table rows
            for (TimesheetRow tr : rows) {
                Row rr = sheet.createRow(r++);
                makeFirstCellBlank.accept(rr);

                boolean nonWorking = "Holiday".equalsIgnoreCase(tr.activity)
                        || "Leave".equalsIgnoreCase(tr.activity)
                        || "WeekOff".equalsIgnoreCase(tr.activity);

                CellStyle rowStyle = wb.createCellStyle();
                rowStyle.cloneStyleFrom(borderDashDot);

                CellStyle shadedStyle = wb.createCellStyle();
                shadedStyle.cloneStyleFrom(borderDashDot);
                shadedStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                shadedStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                CellStyle use = nonWorking ? shadedStyle : rowStyle;
                use.setAlignment(HorizontalAlignment.CENTER);
                use.setVerticalAlignment(VerticalAlignment.CENTER);

                // fill cells from column 1 to 8 (A is blank)
                int col = 1;
                Cell c1 = rr.createCell(col++);
                c1.setCellValue(tr.dayName == null ? "" : tr.dayName);
                c1.setCellStyle(use);

                Cell c2 = rr.createCell(col++);
                c2.setCellValue(tr.date == null ? "" : tr.date);
                c2.setCellStyle(use);

                Cell c3 = rr.createCell(col++);
                c3.setCellValue(tr.inTime == null ? "" : tr.inTime);
                c3.setCellStyle(use);

                Cell c4 = rr.createCell(col++);
                c4.setCellValue(tr.outTime == null ? "" : tr.outTime);
                c4.setCellStyle(use);

                Cell c5 = rr.createCell(col++);
                c5.setCellValue(tr.duration == null ? "" : tr.duration);
                c5.setCellStyle(use);

                Cell c6 = rr.createCell(col++);
                c6.setCellValue(tr.activity == null ? "" : tr.activity);
                c6.setCellStyle(use);

                Cell c7 = rr.createCell(col++);
                c7.setCellValue(tr.managerApproval == null ? "" : tr.managerApproval);
                c7.setCellStyle(use);

                Cell c8 = rr.createCell(col++);
                c8.setCellValue(tr.remarks == null ? "" : tr.remarks);
                c8.setCellStyle(use);
            }

            CellRangeAddress outerBlock3 = new CellRangeAddress(20, r-1, 1, 8);  // B21 → I50

            RegionUtil.setBorderTop(BorderStyle.DOUBLE, outerBlock3, sheet);
            RegionUtil.setBorderBottom(BorderStyle.DOUBLE, outerBlock3, sheet);
            RegionUtil.setBorderLeft(BorderStyle.DOUBLE, outerBlock3, sheet);
            RegionUtil.setBorderRight(BorderStyle.DOUBLE, outerBlock3, sheet);

            // 🔥 Set border color to BLACK
            RegionUtil.setTopBorderColor(IndexedColors.BLACK.getIndex(), outerBlock3, sheet);
            RegionUtil.setBottomBorderColor(IndexedColors.BLACK.getIndex(), outerBlock3, sheet);
            RegionUtil.setLeftBorderColor(IndexedColors.BLACK.getIndex(), outerBlock3, sheet);
            RegionUtil.setRightBorderColor(IndexedColors.BLACK.getIndex(), outerBlock3, sheet);


            // Note and signatures - make sure first cell blank
            //r += 2;
            Row note = sheet.createRow(r++);
            makeFirstCellBlank.accept(note);
            Cell noteCell = note.createCell(1);
            CellRangeAddress range = new CellRangeAddress(r-1, r-1, 1, 8);
            sheet.addMergedRegion(range);
            noteCell.setCellValue("Note: For Regular workday, Inclusive of 30 minutes of lunch break.");

            CellStyle signNDateStyle = wb.createCellStyle();
            headerStyle.setFont(hf);

            r += 4;
            Row sig1 = sheet.createRow(r++);
            makeFirstCellBlank.accept(sig1);
            Cell c1 = sig1.createCell(1);
            Cell c2 = sig1.createCell(6);
            c1.setCellStyle(signNDateStyle);
            c2.setCellStyle(signNDateStyle);
            c1.setCellValue("Customer Signature");
            c2.setCellValue("Employee Signature");


            Row sig1d = sheet.createRow(r++);
            makeFirstCellBlank.accept(sig1d);
            Cell c3 = sig1d.createCell(1);
            Cell c4 = sig1d.createCell(6);
            c3.setCellStyle(signNDateStyle);
            c4.setCellStyle(signNDateStyle);
            c3.setCellValue("Date");
            c4.setCellValue("Date");

            r+=3;
            Row stamp = sheet.createRow(r++);
            makeFirstCellBlank.accept(stamp);
            Cell c5 = stamp.createCell(1);
            Cell c6 = stamp.createCell(6);
            c5.setCellStyle(signNDateStyle);
            c6.setCellStyle(signNDateStyle);
            c5.setCellValue("Customer Stamp");
            c6.setCellValue("Checked by Onsite Project Manager");


            Row sig1dd = sheet.createRow(r++);
            makeFirstCellBlank.accept(sig1d);
            Cell c7=sig1dd.createCell(6);
            c7.setCellStyle(signNDateStyle);
            c7.setCellValue("Date");



            for (int row = sig1.getRowNum()-2; row <= sig1dd.getRowNum(); row++) {   // 5th row is index 4
                Row r1 = sheet.getRow(row);

                if (r1 == null) r1 = sheet.createRow(row);


                for (int col = 1; col <= 8; col++) {  // B → I
                if((row == sig1.getRowNum()+2) && col == 6){

                    break;
                }
                    if(col == 5){continue;}
                    Cell c = r1.getCell(col);
                    if (c == null) c = r1.createCell(col);

                    patchGrayAndDoubleBorder(c, wb);
                }
            }



            for(int row = sig1.getRowNum(); row <= sig1dd.getRowNum(); row++){
                sheet.addMergedRegion(new CellRangeAddress(row,row,1,4));
                sheet.addMergedRegion(new CellRangeAddress(row,row,6,8));
            }



            // Auto-size a few columns
            for (int i = 0; i <= 8; i++) {
                sheet.autoSizeColumn(i);
            }

        // Save the workbook back to disk
        try (FileOutputStream fos = new FileOutputStream(excelFile)) {
            wb.write(fos);
        }

        // Write workbook to bytes
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                wb.write(bos);
                return bos.toByteArray();
        }

    }

    private void patchGrayAndDoubleBorder(Cell cell, Workbook wb) {

        CellStyle newStyle = wb.createCellStyle();

        // clone original style fully
        if (cell.getCellStyle() != null) {
            newStyle.cloneStyleFrom(cell.getCellStyle());
        }

        // Apply shading
        newStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cell.setCellStyle(newStyle);
    }


    // createMergedLabeledRow ensures first cell A is blank always, merges B-C as label and D-I as value
    private static void createMergedLabeledRow(
            Sheet sheet,
            int rowNum,
            String label,
            String value,
            Workbook workbook,
            CellStyle labelStyle,
            CellStyle valueStyle) {

        Row row = sheet.getRow(rowNum);
        if (row == null) row = sheet.createRow(rowNum);

        // Column A always blank
        Cell empty = row.getCell(0);
        if (empty == null) empty = row.createCell(0);
        empty.setCellValue("");

        // Label (B-C)
        Cell l = row.createCell(1);
        l.setCellValue(label == null ? "" : label);
        if (labelStyle != null) l.setCellStyle(labelStyle);

        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 1, 2));

        // Value (D-I)
        Cell v = row.createCell(3);
        v.setCellValue(value == null ? "" : value);
        if (valueStyle != null) v.setCellStyle(valueStyle);

        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 3, 8));
    }

    public byte[] extractSheetsForEmployee() throws Exception {



        XSSFWorkbook sourceWorkbook;

        try (FileInputStream fis = new FileInputStream(excelFile)) {
            sourceWorkbook = new XSSFWorkbook(fis);
        }

        // Create new workbook in memory\
        XSSFWorkbook destWorkbook = new XSSFWorkbook();

        for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); i++) {

            Sheet src = sourceWorkbook.getSheetAt(i);

            String name = src.getSheetName();

            // match Sudhir, Sudip, etc.
            if (name.startsWith(sheetName.split(", ")[0])) {
                Sheet dest = destWorkbook.createSheet(name);
                copySheetFully((XSSFSheet) src, (XSSFSheet) dest, destWorkbook);

                // Auto-size a few columns
                for (int c = 0; c <= 8; c++) {
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


    private void copySheetFully(XSSFSheet src, XSSFSheet dest, XSSFWorkbook destWb) throws Exception {

        // 1️⃣ Column widths
        for (int c = 0; c <= src.getRow(src.getFirstRowNum()).getLastCellNum(); c++) {
            dest.setColumnWidth(c, src.getColumnWidth(c));
        }

        // 2️⃣ Merged regions
        for (int i = 0; i < src.getNumMergedRegions(); i++) {
            dest.addMergedRegion(src.getMergedRegion(i));
        }

        // 3️⃣ Drawings / Images
        copyPictures(src, dest);

        // 4️⃣ Rows + cells
        for (int r = 0; r <= src.getLastRowNum(); r++) {

            XSSFRow srcRow = src.getRow(r);
            if (srcRow == null) continue;

            XSSFRow destRow = dest.createRow(r);
            destRow.setHeight(srcRow.getHeight());

            for (int c = 0; c < srcRow.getLastCellNum(); c++) {

                XSSFCell oldCell = srcRow.getCell(c);
                if (oldCell == null) continue;

                XSSFCell newCell = destRow.createCell(c);

                // Copy style (full clone)
                newCell.setCellStyle(cloneCellStyleFull(destWb, oldCell.getCellStyle()));

                // Copy cell comments
                if (oldCell.getCellComment() != null) {
                    newCell.setCellComment(oldCell.getCellComment());
                }

                // Copy Hyperlink
                if (oldCell.getHyperlink() != null) {
                    newCell.setHyperlink(oldCell.getHyperlink());
                }

                // Copy cell value / formula
                switch (oldCell.getCellType()) {
                    case STRING -> newCell.setCellValue(oldCell.getRichStringCellValue());
                    case NUMERIC -> newCell.setCellValue(oldCell.getNumericCellValue());
                    case BOOLEAN -> newCell.setCellValue(oldCell.getBooleanCellValue());
                    case FORMULA -> newCell.setCellFormula(oldCell.getCellFormula());
                    case BLANK -> newCell.setBlank();
                    default -> {}
                }
            }
        }
    }



    private XSSFCellStyle cloneCellStyleFull(XSSFWorkbook wb, CellStyle old) {

        XSSFCellStyle newStyle = wb.createCellStyle();
        if (old == null) return newStyle;

        newStyle.cloneStyleFrom(old);

        // Need to clone the font too
        XSSFFont oldFont = ((XSSFCellStyle) old).getFont();
        XSSFFont newFont = wb.createFont();
        newFont.setBold(oldFont.getBold());
        newFont.setColor(oldFont.getColor());
        newFont.setFontHeight(oldFont.getFontHeight());
        newFont.setFontName(oldFont.getFontName());
        newFont.setItalic(oldFont.getItalic());
        newFont.setUnderline(oldFont.getUnderline());

        newStyle.setFont(newFont);

        return newStyle;
    }

    private void copyPictures(XSSFSheet src, XSSFSheet dest) throws Exception {

        XSSFDrawing srcDrawing = src.getDrawingPatriarch();
        if (srcDrawing == null) return;

        XSSFDrawing destDrawing = dest.createDrawingPatriarch();

        for (XSSFShape shape : srcDrawing.getShapes()) {

            if (shape instanceof XSSFPicture srcPic) {

                XSSFPictureData picData = srcPic.getPictureData();
                XSSFClientAnchor srcAnchor = (XSSFClientAnchor) srcPic.getAnchor();

                // copy the anchor exactly
                XSSFClientAnchor newAnchor = (XSSFClientAnchor)
                        dest.getWorkbook().getCreationHelper().createClientAnchor();

                newAnchor.setAnchorType(srcAnchor.getAnchorType());
                newAnchor.setCol1(srcAnchor.getCol1());
                newAnchor.setRow1(srcAnchor.getRow1());
                newAnchor.setCol2(srcAnchor.getCol2());
                newAnchor.setRow2(srcAnchor.getRow2());
                newAnchor.setDx1(srcAnchor.getDx1());
                newAnchor.setDy1(srcAnchor.getDy1());
                newAnchor.setDx2(srcAnchor.getDx2());
                newAnchor.setDy2(srcAnchor.getDy2());

                // Add picture
                int picIndex = dest.getWorkbook()
                        .addPicture(picData.getData(), picData.getPictureType());

                XSSFPicture newPic = destDrawing.createPicture(newAnchor, picIndex);

                // 🟢 NO resize() here — keeps original size + position exactly
            }
        }
    }



}
