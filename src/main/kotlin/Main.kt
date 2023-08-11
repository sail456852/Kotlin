import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.json.JSONArray
import java.io.File
import java.io.FileOutputStream

fun main() {
    val jsonFile = "/Users/eugene/IdeaProjects/Gradle/src/main/resources/input.json"
    val excelFile = "output.xlsx"

    // Read JSON file
    val jsonString = File(jsonFile).readText()
    val jsonArray = JSONArray(jsonString)

    // Create a new Excel workbook
    val workbook = XSSFWorkbook()
    val sheet = workbook.createSheet("Data")

    // Create header row
    val headerRow = sheet.createRow(0)
    val headers = arrayOf("Name", "Age", "Email")
    for (i in headers.indices) {
        val cell = headerRow.createCell(i)
        cell.setCellValue(headers[i])
    }

    // Populate data rows
    for (i in 0 until jsonArray.length()) {
        val jsonObject = jsonArray.getJSONObject(i)
        val dataRow = sheet.createRow(i + 1)

        val name = jsonObject.getString("name")
        val age = jsonObject.getInt("age")
        val email = jsonObject.getString("email")

        val nameCell = dataRow.createCell(0)
        nameCell.setCellValue(name)

        val ageCell = dataRow.createCell(1)
        ageCell.setCellValue(age.toDouble())

        val emailCell = dataRow.createCell(2)
        emailCell.setCellValue(email)
    }

    // Auto-size columns
    for (i in headers.indices) {
        sheet.autoSizeColumn(i)
    }

    // Save Excel file
    val fileOut = FileOutputStream(excelFile)
    workbook.write(fileOut)
    fileOut.close()

    // Close the workbook
    workbook.close()

    println("Excel file generated successfully.")
}