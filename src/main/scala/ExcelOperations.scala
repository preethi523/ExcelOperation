import com.typesafe.scalalogging._
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.xssf.usermodel.{XSSFRow, XSSFSheet, XSSFWorkbook}

import java.io._
import java.nio.file.{Files, Paths}

object ExcelOperations {
  val logger: Logger = Logger("log")

  def main(args: Array[String]): Unit = {
    val path ="src/main/excelSheets/Excel.xlsx"
    val arr = Array(Array(1, "Teja"), Array(22, "Teja"))
    createExcel(path)
    println(writeIntoSingleExcelCell(path, 4, 0, "3"))
    println(writeIntoMultipleCell(path, Array(3, 2), Array(1, 0), arr))
    println(updateDataInExcel(path, 3, 1, "Preethi Anbu"))
    readSingleCellValue(path, 3, 0)
    readMultipleCellValues(path, Array(3, 2), Array(1, 0))
    println(searchIfExist(path, "Teja"))
    println(searchAndReplaceFirstOccurance(path, "Teja", "teja anbhu"))
    print(deleteDataInExcel(path, 3, 1))
    print(deleteExcel(path))

  }

  /** *
   * creates an excel workbook for the project at specified path
   *@param path: path where Excel should be created
   * */
  def createExcel(path: String): String = {
    ExcelOperations.logger.info("createExcel Method started")
    if(!Files.exists(Paths.get(path))){           // check if the path is present
      val fileName = Paths.get(path).getFileName
      val extension = fileName.toString.split("\\.").last  // to find the file extension
      if (extension == "xlsx" || extension == "xls") {
        val workbook: XSSFWorkbook = new XSSFWorkbook
        val spreadsheet: XSSFSheet = workbook.createSheet("Student_Data")
        val out = new FileOutputStream(new File(path))
        workbook.write(out)
        logger.info("Excel File has been created successfully.")
        return "Excel File has been created successfully."
      }
      "Not an excel"
    }
    else {
      logger.error("File already exist in the specified path.")
      "File already exist in the specified path."
    }
  }

  /** *
   * write into single cell of an excel
   *
   * @param path :location where file is created
   * @param row  :Row of excel to write
   * @param col  :column of excel to write
   * @param data :data to be written in cell
   */
  def writeIntoSingleExcelCell(path: String, row: Int, col: Int, data: String): String = {
    logger.info("writeIntoSingleExcelCell Method started")
    var fileName: String = null
    var fileOut: FileOutputStream = null
    try {
      fileName = fileName + Paths.get(path).getFileName
      fileOut = new FileOutputStream(path)
    }
    catch {
      case e1: FileNotFoundException =>
        logger.error("Given path is incorrect" + e1)
        return "Given path is incorrect"
    }
    val extension = fileName.split("\\.").last
    if (extension == "xlsx" || extension == "xls") {
      val workBook = new XSSFWorkbook()
      //invoking createSheet() method and passing the name of the sheet to be created
      val sheet = workBook.createSheet("studentData")
      //creating  excelRow using the createRow() method
      val excelRow: XSSFRow = sheet.createRow(row)
      //creating cell by using the createCell() method and setting the values using the setCellValue() method
      excelRow.createCell(col).setCellValue(data)
      // val fileOut = new FileOutputStream(path)
      workBook.write(fileOut)
      fileOut.close()
      workBook.close()
      """data written"""
    }
    else "Not a excel file"

  }

  /**
   * Delete a particular position in excel
   *
   * @param path :location of file
   * @param row  :excelRow of excel
   * @param col  :column of excel
   */
  def deleteDataInExcel(path: String, row: Int, col: Int): String = {
    logger.info("deleteDataInExcel Method started")
    try {
      val fileInputStream = new FileInputStream(new File(path))
      val workBook = new XSSFWorkbook(fileInputStream)
      val sheet = workBook.getSheetAt(0)
      val excelRow = sheet.getRow(row)
      val cell = excelRow.createCell(col)
      excelRow.removeCell(cell)              // remove the needed cell
      val outFile = new FileOutputStream(new File(path))
      workBook.write(outFile)
      workBook.close()
      outFile.close()
      "deleted the data at cell"
    } catch {
      case e1: FileNotFoundException =>
        logger.error("File not found in the specified path")
        "File not found in the specified path"
      case e2: NullPointerException =>
        logger.error("Cell does not contain any value")
        "Cell does not contain any value"
      case e3: IllegalArgumentException =>
        logger.error("Invalid Sheet number")
        "Invalid Sheet number"

    }
  }

  /** *
   * delete an Excel in given path
   *
   * @param path :Location of the given file
   * */
  def deleteExcel(path: String): String = {
    ExcelOperations.logger.info(" deleteExcel Method started")
    if (!Files.exists(Paths.get(path))) {
      ExcelOperations.logger.info("function ended")
      "file not found"
    }
    else {
      val file = new File(path)
      file.delete()
      ExcelOperations.logger.info("path deleted")
      "path deleted"
    }
  }

  /** *
   * Search and replace the given string with another string at its first appearance in Excel sheet
   *
   * @param path    :Location of the file
   * @param search  :String to be searched
   * @param replace :String to be replaced
   */
  def searchAndReplaceFirstOccurance(path: String, search: String, replace: String): String = {
    ExcelOperations.logger.info("searchAndReplaceFirstOccurance Method started")
    var count = 0
    try {
      val fin = new FileInputStream(new File(path))
      val workbook = new XSSFWorkbook(fin)
      val mySheet = workbook.getSheetAt(0)
      val rowIterator = mySheet.iterator()
      rowIterator.hasNext
    } catch {
      case e1: FileNotFoundException =>
        return "File not found in the specified path"
      case e3: IllegalArgumentException =>
        return "Invalid Sheet number"
    }
    val fin = new FileInputStream(new File(path))
    val workbook = new XSSFWorkbook(fin)
    val mySheet = workbook.getSheetAt(0)
    val rowIterator = mySheet.iterator()
    while (rowIterator.hasNext) {
      val row = rowIterator.next()
      val cellIterator = row.cellIterator()
      while (cellIterator.hasNext) {
        val cell = cellIterator.next()
        if ((cell.getStringCellValue == search) && (count == 0)) {
          cell.setCellValue(replace)
          count = count + 1
          val outFile = new FileOutputStream(new File(path))
          workbook.write(outFile)
        }
      }
    }
    if (count >= 1) "found and replaced"
    else "not found"
  }

  /** *
   * to write into multiple cell.
   *
   * @param path :Location where file to be found
   * @param row  :Array of excelRow values
   * @param col  :Array of column values
   * @param arr  :Array[Array[Any]] to write data in square matrix form
   */
  def writeIntoMultipleCell(path: String, row: Array[Int], col: Array[Int], arr: Array[Array[Any]]): String = {
    ExcelOperations.logger.info(" writeIntoMultipleCell Method started")
    var fileName: String = null
    var fileOut : FileOutputStream = null
    try {
      fileName = fileName + Paths.get(path).getFileName
      fileOut = new FileOutputStream(new File(path))
    }
    catch {
      case e1: FileNotFoundException =>
        logger.error("Given path is incorrect")
        return "Given path is incorrect"
    }
    val extension = fileName.split("\\.").last
    if (matrix(arr) && (extension == "xlsx" || extension == "xls") && (row.length == arr.length) && (arr.length == col.length)) {
      //checks whether the input array is of square -matrix form  and row and col array must be same length as arr
      val workBook = new XSSFWorkbook()
      val sheet = workBook.createSheet("Student_Data")
      for (i <- row.indices) {
        val excelRow = sheet.createRow(row(i))
        for (c <- col.indices) {
          val excelCell = excelRow.createCell(col(c))
          val value = arr(i)(c).toString
          excelCell.setCellValue(value)
        }
      }
      val fileOut = new FileOutputStream(new File(path))
      workBook.write(fileOut)
      fileOut.close()
      "wrote data into multiple cell"
    }
    else "Not an excel file or input incorrect"
  }


  /** *
   *
   * Checks  whether given input array is of form of  an square matrix
   *
   * @param arr :The Array[Array[]] to be checked
   * @return Boolean true/false
   */
  def matrix(arr: Array[Array[Any]]): Boolean = {
    for (i <- arr.indices) {
      if (arr(i).length != arr.length) {
        return false
      }
    }
    true
  }

  /** *
   * reads multiple cell values
   *
   * @param path :location where file is available to read
   * @param row  :array of excelRow values
   * @param col  :array of column values
   */

  def readMultipleCellValues(path: String, row: Array[Int], col: Array[Int]): Any = {
    ExcelOperations.logger.info("readMultipleCellValues Method started")
    try {
      val file = new FileInputStream(new File(path))
      val workbook = new XSSFWorkbook(file)
      if (row.length == col.length) {
        val mySheet = workbook.getSheetAt(0)
        for (i <- row; j <- col) {
          val excelRow = mySheet.getRow(i)
          val excelCol = excelRow.getCell(j)
          if (excelCol.getCellType == Cell.CELL_TYPE_STRING) {
            val result: String = excelCol.getStringCellValue
            ExcelOperations.logger.info(result)
          }
          else excelCol.getCellType == Cell.CELL_TYPE_BLANK
          print(" ")
        }
        println(" ")
        workbook.close()
        ExcelOperations.logger.info("Read data")
        "Read data"
      }
    }

    catch {
      case e1: FileNotFoundException =>
        logger.error(e1+"File not found in the specified path")
        "File not found in the specified path"
      case e2: IllegalArgumentException =>
        logger.error(e2+"Invalid Sheet number")
        "Invalid Sheet number"
      case e3: NullPointerException =>
        logger.error(e3 + "The cells are empty")
        "The cells are empty"
    }
  }


  /** *
   * reads a single cell value
   *
   * @param path :Location to read a file
   * @param row  :Row to read a cell
   * @param col  :column to read a column
   */

  def readSingleCellValue(path: String, row: Int, col: Int): String = {
    ExcelOperations.logger.info(" readSingleCellValue Method started")
    try {
      val fileInputStream = new FileInputStream(new File(path))
      val workBook = new XSSFWorkbook(fileInputStream)
      val sheet = workBook.getSheetAt(0)
      val excelRow = sheet.getRow(row)
      val value = excelRow.getCell(col)
      val result = value.getStringCellValue
      ExcelOperations.logger.info(result)
      workBook.close()
      ExcelOperations.logger.info("Read data")
      "Read data"
    }
    catch {
      case e1: FileNotFoundException =>
        logger.error("File not found in the specified path")
        "File not found in the specified path"
      case e2: NullPointerException =>
        logger.error("Cell does not contain any value")
        "Cell does not contain any value"
      case e3: IllegalArgumentException =>
        logger.error("Invalid Sheet number")
        "Invalid Sheet number"
    }
  }

  /** *
   * updates data into excel at particular position
   *
   * @param path :Location where file is present
   * @param row  :excelRow of an excel
   * @param col  :column of excel
   * @param data :Data to be updated
   */

  def updateDataInExcel(path: String, row: Int, col: Int, data: String): String = {
    ExcelOperations.logger.info("updateDataInExcel Method started")
    try {
      val fileInputStream = new FileInputStream(new File(path))
      val workBook = new XSSFWorkbook(fileInputStream)
      val sheet = workBook.getSheetAt(0)
      val excelRow = sheet.getRow(row)
      excelRow.createCell(col).setCellValue(data)  // creates empty cell and set new value in cell
      val fileOut = new FileOutputStream(new File(path))
      workBook.write(fileOut)
      fileOut.close()
      workBook.close()
      "Updated data"
    }
    catch {
      case e1: FileNotFoundException =>
        logger.error("File not found in the specified path")
        "File not found in the specified path"
      case e2: NullPointerException =>
        logger.error("Cell does not contain any value")
        "Cell does not contain any value"
      case e3: IllegalArgumentException =>
        logger.error("Invalid Sheet number")
        "Invalid Sheet number"

    }
  }

  /** *
   * searches in excel sheet whether the given input string is present or not
   *
   * @param path   :Location of the file
   * @param search :String to be found
   * */

  def searchIfExist(path: String, search: String): String = {
    ExcelOperations.logger.info("searchIfExist method started")
    var result = false
    try {
      val fin = new FileInputStream(new File(path))
      val workbook = new XSSFWorkbook(fin)
      val mySheet = workbook.getSheetAt(0)
      val rowIterator = mySheet.iterator()
      rowIterator.hasNext
    }
    catch {
      case e1: FileNotFoundException =>
        return "File not found in the specified path"
      case e3: IllegalArgumentException =>
        return "Invalid Sheet number"
      case e: NullPointerException =>
        return "cells are empty"
    }
    val fin = new FileInputStream(new File(path))
    val workbook = new XSSFWorkbook(fin)
    val mySheet = workbook.getSheetAt(0)
    val rowIterator = mySheet.iterator()
    while (rowIterator.hasNext) {
      val row = rowIterator.next()
      val cellIterator = row.cellIterator()
      while (cellIterator.hasNext) {
        val cell = cellIterator.next()
        if (cell.getStringCellValue == search) {
          result = true
        }
      }
    }
    ExcelOperations.logger.info("Searched")
    if (result)
      "found"
    else
      "Data not found"
  }


}
