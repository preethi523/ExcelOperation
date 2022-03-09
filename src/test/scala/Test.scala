import org.scalatest.funsuite.AnyFunSuite

class Test extends AnyFunSuite {
  //test input data
  val correctPath = "src/main/excelSheets/Excel.xlsx"
  // should be present
  val incorrectPath = "xyzxx/pr.xlsx"
  val createxlsx= "src/main/excelSheets/Excel1.xlsx"
  //should not be present
  val path ="/src/main/excelSheets/pr.xlsx"
  val Excelpath="jjjj.xlsx"
  //not be present
  val pathTxt = "src/main/excelSheets/hi.txt"
  val pathTxtCreate = "src/main/excelSheets/preethi.txt"
  //not be present

  val arr = Array(Array(1, "Teja"), Array(22, "Teja")) //input to write into multiple cell

  test("create Excel") {
 assert(ExcelOperations.createExcel(correctPath) == "File already exist in the specified path.")
    assert(ExcelOperations.createExcel(createxlsx) == "Excel File has been created successfully.")
   assert(ExcelOperations.createExcel(pathTxtCreate) == "Not an excel")
  }

  test("write into Excel") {
    assert(ExcelOperations.writeIntoSingleExcelCell(correctPath, 2, 3, "hi") == "data written")
    assert(ExcelOperations.writeIntoSingleExcelCell(pathTxt, 2, 3, "hi") == "Not a excel file")
    assert(ExcelOperations.writeIntoSingleExcelCell(incorrectPath, 2, 3, "hoc") == "Given path is incorrect")
  }

  test("write multiple data ") {
    assert(ExcelOperations.writeIntoMultipleCell(correctPath, Array(3, 2), Array(0, 1), arr) == "wrote data into multiple cell")
    assert(ExcelOperations.writeIntoMultipleCell(incorrectPath, Array(3, 2), Array(0, 1), arr) == "Given path is incorrect")
    assert(ExcelOperations.writeIntoMultipleCell(pathTxt, Array(3, 2), Array(0, 1), arr) == "Not an excel file or input incorrect")
  }
  test("read multiple data ") {
    assert(ExcelOperations.readMultipleCellValues(createxlsx,Array(3, 2), Array(0, 1))=="The cells are empty")
    assert(ExcelOperations.readMultipleCellValues(incorrectPath, Array(3, 2), Array(0, 1)) == "File not found in the specified path")
    assert(ExcelOperations.readMultipleCellValues(correctPath,Array(3, 2), Array(0, 1))=="Read data")
  }
  test("read single data ") {
    assert(ExcelOperations.readSingleCellValue(createxlsx,3, 1)=="Cell does not contain any value")
    assert(ExcelOperations.readSingleCellValue(incorrectPath, 3, 1) == "File not found in the specified path")
    assert(ExcelOperations.readSingleCellValue(correctPath,3 ,1)=="Read data")
  }

  test("update multiple data") {
    assert(ExcelOperations.updateDataInExcel(correctPath, 0, 5, "hi") == "Cell does not contain any value")
    assert(ExcelOperations.updateDataInExcel(correctPath, 3, 1, "hi") == "Updated data")
  }

  test("search in Excel") {
    assert(ExcelOperations.searchIfExist(correctPath, "Teja") == "found")
    assert(ExcelOperations.searchIfExist(correctPath, "teja nn") == "Data not found")
    assert(ExcelOperations.searchIfExist(incorrectPath, "teja") == "File not found in the specified path")
  }
  test("search and replace in Excel") {
   assert(ExcelOperations.searchAndReplaceFirstOccurance(correctPath, "Teja", "11") == "found and replaced")
    assert(ExcelOperations.searchAndReplaceFirstOccurance(correctPath, "hi", "teja") =="found and replaced")
  }
  test("delete data in excel") {
    assert(ExcelOperations.deleteDataInExcel(correctPath, 3, 1) == "deleted the data at cell")
    assert(ExcelOperations.deleteDataInExcel(correctPath, 0, 5) == "Cell does not contain any value")
  }

  test("delete Excel") {
   assert(ExcelOperations.deleteExcel(path) == "file not found")
  assert(ExcelOperations.deleteExcel(createxlsx) == "path deleted" )
  }

}
