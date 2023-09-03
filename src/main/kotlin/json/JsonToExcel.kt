package json

import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.io.FileOutputStream

fun main() {
    val excelFile = "output.xlsx"
    val jsonArray = readJson()

    val workbook = XSSFWorkbook()
    var sheet = workbook.createSheet("Data")

    val headers = arrayOf(
        "City Name",
        "City Latitude",
        "City Longitude",
        "Station Name",
        "Station Latitude",
        "Station Longitude"
    )

    sheet = createHeaderRow(
        headers, sheet
    )

    for (i in 0 until jsonArray.length()) {
        val jsonObject = jsonArray.getJSONObject(i)
        val stations = jsonObject.getJSONArray("Stations")
        populateByRow(stations, sheet, headers, jsonObject)
    }

    val fileOut = FileOutputStream(excelFile)
    fileOut.use {
        workbook.use {
            workbook.write(fileOut)
        }
    }

    println("Excel file generated successfully.")
}

private fun populateByRow(
    stations: JSONArray,
    sheet: XSSFSheet,
    headers: Array<String>,
    jsonObject: JSONObject
) {
    var currentRowIndex1 = 1
    (0 until stations.length()).forEach { stationIndex ->
        val dataRow = sheet.createRow(currentRowIndex1)
        currentRowIndex1++
        headers.indices.forEach { columnIndex ->
            when (columnIndex) {
                0 -> {
                    sheet.autoSizeColumn(columnIndex)
                    val cell = dataRow.createCell(columnIndex)
                    val cityName = jsonObject.getJSONObject("Place")["name"]
                    cell.setCellValue(cityName.toString())
                }

                1 -> {
                    sheet.autoSizeColumn(columnIndex)
                    val cell = dataRow.createCell(columnIndex)
                    val cityLat = jsonObject.getJSONObject("Place").getJSONArray("geo")[0]
                    cell.setCellValue(cityLat.toString())
                }

                2 -> {
                    sheet.autoSizeColumn(columnIndex)
                    val cell = dataRow.createCell(columnIndex)
                    val cityLog = jsonObject.getJSONObject("Place").getJSONArray("geo")[1]
                    cell.setCellValue(cityLog.toString())
                }

                3 -> {
                    sheet.autoSizeColumn(columnIndex)
                    val cell = dataRow.createCell(columnIndex)
                    val stationName = stations.getJSONObject(stationIndex)["Name"]
                    cell.setCellValue(stationName.toString())
                }

                4 -> {
                    sheet.autoSizeColumn(columnIndex)
                    val cell = dataRow.createCell(columnIndex)
                    val stationLat = stations.getJSONObject(stationIndex)["Latitude"]
                    cell.setCellValue(stationLat.toString())
                }

                5 -> {
                    sheet.autoSizeColumn(columnIndex)
                    val cell = dataRow.createCell(columnIndex)
                    val stationLongitude = stations.getJSONObject(stationIndex)["Longitude"]
                    cell.setCellValue(stationLongitude.toString())
                }
            }
        }
    }
}

private fun readJson(): JSONArray {
    val jsonFile = "/Users/eugene/IdeaProjects/Gradle/src/main/resources/input.json"
    val jsonString = File(jsonFile).readText()
    return JSONArray(jsonString)
}

private fun createHeaderRow(headers: Array<String>, sheet: XSSFSheet): XSSFSheet {
    val headerRow = sheet.createRow(0)
    for (i in headers.indices) {
        val cell = headerRow.createCell(i)
        cell.setCellValue(headers[i])
    }
    return sheet
}