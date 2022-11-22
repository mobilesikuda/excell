import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream

 fun main(args: Array<String>) {
     val filepath = "./test_file.xlsx"
     val file = File(filepath)

     //Instantiate Excel workbook:
     val xlWb = XSSFWorkbook()
     //Instantiate Excel worksheet:
     val xlWs = xlWb.createSheet()
     //Row index specifies the row in the worksheet (starting at 0):
     val rowNumber = 0
     //Cell index specifies the column within the chosen row (starting at 0):
     val columnNumber = 0
     //Write text value to cell located at ROW_NUMBER / COLUMN_NUMBER:
     val xlRow = xlWs.createRow(rowNumber)
     val xlCol = xlRow.createCell(columnNumber)
     xlCol.setCellValue("Chercher Tech")
     //Write file:
     val outputStream = FileOutputStream(filepath)
     xlWb.write(outputStream)
     xlWb.close()
 }