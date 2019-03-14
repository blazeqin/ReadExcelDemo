package com.lion.controller

import com.lion.model.ExcelInfo
import com.lion.util.JsonResult
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.web.bind.annotation.*
import org.springframework.web.multipart.MultipartFile
import java.io.*
import java.text.DecimalFormat
import java.text.SimpleDateFormat
import java.util.*

@RestController
@RequestMapping
class ExcelProcessController {

    private val df = DecimalFormat("0")
    private val sdf = SimpleDateFormat("yyyy-MM-dd")

    @PostMapping("/parseExcel")
    fun parseExcel(@RequestParam("file") file: MultipartFile): Any {
        val result = mutableMapOf<String, Any>()
        // 获得上传文件的文件名
        val fileName = file.originalFilename
        // 获取文件扩展名
        val eName = fileName.substring(fileName.lastIndexOf(".") + 1)
        var inputStream: InputStream? = null
        try {
            inputStream = file.inputStream
            val workbook = getWorkbook(inputStream, eName)
            val dataList = getExcelContent(workbook, result)

            result.put("result", dataList)
        } catch (e: IOException) {
            e.printStackTrace()
        } catch (e: NumberFormatException) {
            e.printStackTrace()
        } finally {
            closeStream(inputStream)
        }
        return JsonResult.ok(result);
    }

    private fun getExcelContent(workbook: Workbook?, result: MutableMap<String, Any>): ArrayList<ExcelInfo> {
        // 获取工作薄第一张表
        val sheet = workbook!!.getSheetAt(0)
        // 获取第一行
        var row: Row? = sheet.getRow(0)
        // 获取有效单元格数
        val cellNum = row!!.physicalNumberOfCells
        // 表头集合
        val headList = ArrayList<String>()
        for (i in 0 until cellNum) {
            val cell = row.getCell(i)
            val data = cell.stringCellValue
            headList.add(data)
        }

        // 获得有效行数
        val rowNum = sheet.getPhysicalNumberOfRows()
        result.put("sum", rowNum - 1)
        val dataList = ArrayList<ExcelInfo>()
        val rowIterator = sheet.rowIterator()
        //去掉头部数据
        rowIterator.next()
        var data: ExcelInfo? = null
        for (i in 1 until rowNum) {
            //row = sheet.getRow(i);//如果有空行，Workbook读取excel数据为null
            row = rowIterator.next()
            if (row != null) {
                data = ExcelInfo()
                for (j in headList.indices) {
                    // 解析单元格
                    val cellData = getCellFormatValue(row!!.getCell(j)).toString()
                    // 根据字段给字段设值
                    saveCellDataToBean(headList, data, j, cellData)
                }
                dataList.add(data)
            }
        }
        return dataList
    }

    private fun saveCellDataToBean(headList: List<String>, data: ExcelInfo, j: Int, cellData: String) {
        when (headList[j]) {
            "string" -> data.string = cellData
            "gender"
            -> {
                var gender = 0
                if (cellData.isNotEmpty()) {
                    try {
                        gender = cellData.toInt()
                        if (gender > 2 || gender < 0) gender = 0
                    } catch (e: NumberFormatException) {
                        if (cellData == "男") {
                            gender = 1
                        } else if (cellData == "女") {
                            gender = 2
                        }
                    }

                }
                data.gender = gender
            }
            "int" -> if (cellData.isNullOrEmpty()) {
                data.int = 0
            } else {
                data.int = cellData.toInt()
            }
            else -> {
            }
        }
    }

    /*
     * 根据excel文件格式获知excel版本信息
     */
    private fun getWorkbook(fs: InputStream, str: String): Workbook? {
        var book: Workbook? = null
        try {
            if ("xls" == str) {
                // 2003
                book = HSSFWorkbook(fs)
            } else {
                // 2007以上
                book = XSSFWorkbook(fs)
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }

        return book
    }

    /**
     * 获取单个单元格数据
     */
    private fun getCellFormatValue(cell: Cell?): Any {
        val cellValue: Any
        if (cell != null) {
            // 判断cell类型
            when (cell.cellType) {
                CellType.NUMERIC -> {
                    cellValue = df.format(cell.numericCellValue)
                }
                CellType.FORMULA -> {
                    // 判断cell是否为日期格式
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // 转换为日期格式YYYY-mm-dd
                        cellValue = cell.dateCellValue
                    } else {
                        // 数字
                        cellValue = df.format(cell.numericCellValue)
                    }
                }
                CellType.STRING -> {
                    cellValue = cell.richStringCellValue.string
                }
                else -> cellValue = ""
            }
        } else {
            cellValue = ""
        }
        return cellValue
    }

    private fun closeStream(stream: Closeable?) {
        try {
            stream?.close()
        } catch (e: IOException) {
            e.printStackTrace()
        }

    }
}