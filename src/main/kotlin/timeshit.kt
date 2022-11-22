/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.apache.poi.examples.ss

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream

/**
 * A weekly timesheet created using Apache POI.
 * Usage:
 * TimesheetDemo -xls|xlsx
 */
object TimesheetDemo {
    private val titles = arrayOf(
        "Person", "ID", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun",
        "Total\nHrs", "Overtime\nHrs", "Regular\nHrs"
    )
    private val sample_data = arrayOf(
        arrayOf<Any?>("Yegor Kozlov", "YK", 5.0, 8.0, 10.0, 5.0, 5.0, 7.0, 6.0),
        arrayOf<Any?>("Gisella Bronzetti", "GB", 4.0, 3.0, 1.0, 3.5, null, null, 4.0)
    )

    @Throws(Exception::class)
    @JvmStatic
    fun main(args: Array<String>) {
        val wb: Workbook
        wb = if (args.size > 0 && args[0] == "-xls") HSSFWorkbook() else XSSFWorkbook()
        val styles = createStyles(wb)
        val sheet = wb.createSheet("Timesheet")
        val printSetup = sheet.printSetup
        printSetup.landscape = true
        sheet.fitToPage = true
        sheet.horizontallyCenter = true

        //title row
        val titleRow = sheet.createRow(0)
        titleRow.heightInPoints = 45f
        val titleCell = titleRow.createCell(0)
        titleCell.setCellValue("Weekly Timesheet")
        titleCell.cellStyle = styles["title"]
        sheet.addMergedRegion(CellRangeAddress.valueOf("\$A$1:\$L$1"))

        //header row
        val headerRow = sheet.createRow(1)
        headerRow.heightInPoints = 40f
        var headerCell: Cell
        for (i in titles.indices) {
            headerCell = headerRow.createCell(i)
            headerCell.setCellValue(titles[i])
            headerCell.cellStyle = styles["header"]
        }
        var rownum = 2
        for (i in 0..9) {
            val row = sheet.createRow(rownum++)
            for (j in titles.indices) {
                val cell = row.createCell(j)
                if (j == 9) {
                    //the 10th cell contains sum over week days, e.g. SUM(C3:I3)
                    val ref = "C$rownum:I$rownum"
                    cell.cellFormula = "SUM($ref)"
                    cell.cellStyle = styles["formula"]
                } else if (j == 11) {
                    cell.cellFormula = "J$rownum-K$rownum"
                    cell.cellStyle = styles["formula"]
                } else {
                    cell.cellStyle = styles["cell"]
                }
            }
        }

        //row with totals below
        var sumRow = sheet.createRow(rownum++)
        sumRow.heightInPoints = 35f
        var cell: Cell
        cell = sumRow.createCell(0)
        cell.cellStyle = styles["formula"]
        cell = sumRow.createCell(1)
        cell.setCellValue("Total Hrs:")
        cell.cellStyle = styles["formula"]
        for (j in 2..11) {
            cell = sumRow.createCell(j)
            val ref = ('A'.code + j).toChar().toString() + "3:" + ('A'.code + j).toChar() + "12"
            cell.cellFormula = "SUM($ref)"
            if (j >= 9) cell.cellStyle = styles["formula_2"] else cell.cellStyle = styles["formula"]
        }
        rownum++
        sumRow = sheet.createRow(rownum++)
        sumRow.heightInPoints = 25f
        cell = sumRow.createCell(0)
        cell.setCellValue("Total Regular Hours")
        cell.cellStyle = styles["formula"]
        cell = sumRow.createCell(1)
        cell.cellFormula = "L13"
        cell.cellStyle = styles["formula_2"]
        sumRow = sheet.createRow(rownum++)
        sumRow.heightInPoints = 25f
        cell = sumRow.createCell(0)
        cell.setCellValue("Total Overtime Hours")
        cell.cellStyle = styles["formula"]
        cell = sumRow.createCell(1)
        cell.cellFormula = "K13"
        cell.cellStyle = styles["formula_2"]

        //set sample data
        for (i in sample_data.indices) {
            val row = sheet.getRow(2 + i)
            for (j in sample_data[i].indices) {
                if (sample_data[i][j] == null) continue
                if (sample_data[i][j] is String) {
                    row.getCell(j).setCellValue(sample_data[i][j] as String?)
                } else {
                    row.getCell(j).setCellValue((sample_data[i][j] as Double?)!!)
                }
            }
        }

        //finally set column widths, the width is measured in units of 1/256th of a character width
        sheet.setColumnWidth(0, 30 * 256) //30 characters wide
        for (i in 2..8) {
            sheet.setColumnWidth(i, 6 * 256) //6 characters wide
        }
        sheet.setColumnWidth(10, 10 * 256) //10 characters wide

        // Write the output to a file
        var file = "timesheet.xls"
        if (wb is XSSFWorkbook) file += "x"
        val out = FileOutputStream(file)
        wb.write(out)
        out.close()
    }

    /**
     * Create a library of cell styles
     */
    private fun createStyles(wb: Workbook): Map<String, CellStyle> {
        val styles: MutableMap<String, CellStyle> = HashMap()
        var style: CellStyle
        val titleFont = wb.createFont()
        titleFont.fontHeightInPoints = 18.toShort()
        titleFont.bold = true
        style = wb.createCellStyle()
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        style.setFont(titleFont)
        styles["title"] = style
        val monthFont = wb.createFont()
        monthFont.fontHeightInPoints = 11.toShort()
        monthFont.color = IndexedColors.WHITE.getIndex()
        style = wb.createCellStyle()
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        style.fillForegroundColor = IndexedColors.GREY_50_PERCENT.getIndex()
        style.fillPattern = FillPatternType.SOLID_FOREGROUND
        style.setFont(monthFont)
        style.wrapText = true
        styles["header"] = style
        style = wb.createCellStyle()
        style.alignment = HorizontalAlignment.CENTER
        style.wrapText = true
        style.borderRight = BorderStyle.THIN
        style.rightBorderColor = IndexedColors.BLACK.getIndex()
        style.borderLeft = BorderStyle.THIN
        style.leftBorderColor = IndexedColors.BLACK.getIndex()
        style.borderTop = BorderStyle.THIN
        style.topBorderColor = IndexedColors.BLACK.getIndex()
        style.borderBottom = BorderStyle.THIN
        style.bottomBorderColor = IndexedColors.BLACK.getIndex()
        styles["cell"] = style
        style = wb.createCellStyle()
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        style.fillForegroundColor = IndexedColors.GREY_25_PERCENT.getIndex()
        style.fillPattern = FillPatternType.SOLID_FOREGROUND
        style.dataFormat = wb.createDataFormat().getFormat("0.00")
        styles["formula"] = style
        style = wb.createCellStyle()
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        style.fillForegroundColor = IndexedColors.GREY_40_PERCENT.getIndex()
        style.fillPattern = FillPatternType.SOLID_FOREGROUND
        style.dataFormat = wb.createDataFormat().getFormat("0.00")
        styles["formula_2"] = style
        return styles
    }
}