package com.github.saisse.excel_scala

import java.io.{File, FileInputStream, InputStream}
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.{Sheet => PoiSheet, Workbook => PoiWorkbook, Cell => PoiCell}
import org.apache.poi.ss.format.CellFormat
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.joda.time.DateTime

class Book(book: PoiWorkbook) {
  def sheet(name: String): Sheet = new Sheet(book.getSheet(name))
}

class Sheet(sheet: PoiSheet) {
  def listRows(startRow: Int, end: (Int) => Boolean): Seq[Int] = {
    end(startRow) match {
      case false => Nil
      case true => startRow +: listRows(startRow + 1, end)
    }
  }

  def parseRow[A](row: Int)(parser: (Int) => A): A = parser(row)
  def parseRows[A](startRow: Int, end: (Int) => Boolean)(parser: (Int) => A): Seq[A] = {
    listRows(startRow, end).map(row => parser(row))
  }

  def cell(column: String, row: Int): String = s"$column$row"

  def string(cell: String): String = string(Cell.split(cell))
  def stringOpt(cell: String): Option[String] = {
    val address = Cell.split(cell)

    opt(address, (address: (Int, Int)) => {
      string(address) match {
      case "" => None
      case s => Some(s)
    }
    })
  }
  private def string(address: (Int, Int)): String = {
    val c = poiCell(address)

    c.getCellType() match {
      case PoiCell.CELL_TYPE_STRING => c.getRichStringCellValue.getString
      case PoiCell.CELL_TYPE_FORMULA => c.getRichStringCellValue.getString
      case PoiCell.CELL_TYPE_BLANK => ""
      case PoiCell.CELL_TYPE_NUMERIC => {
        val cf = CellFormat.getInstance(c.getCellStyle.getDataFormatString)
        cf.apply(c).text
      }
    }
  }

  private def opt[V](address: (Int, Int), toValue: ((Int, Int)) => Option[V]): Option[V] = {
    exists(address) match {
      case true => toValue(address)
      case false => None
    }
  }

  def print(cell: String): Unit = {
    val c = poiCell(Cell.split(cell))
    println(s"${c.getCellType}")
    println(s"${c.getCellStyle.getDataFormat}")
    println(s"${c.getCellStyle.getDataFormatString}")
  }

  def double(cell: String): Double = double(Cell.split(cell))
  def doubleOpt(cell: String): Option[Double] = opt(Cell.split(cell), (a: (Int, Int)) => Some(double(a)))
  private def double(address: (Int, Int)): Double = poiCell(address).getNumericCellValue

  def int(cell: String): Int = int(Cell.split(cell))
  def intOpt(cell: String): Option[Int] = opt(Cell.split(cell), (a: (Int, Int)) => Some(int(a)))
  private def int(address: (Int, Int)): Int = poiCell(address).getNumericCellValue.toInt

  def dateTime(cell: String): DateTime = dateTime(Cell.split(cell))
  def dateTimeOpt(cell: String): Option[DateTime] = opt(Cell.split(cell), (a: (Int, Int)) => Some(dateTime(a)))
  private def dateTime(address: (Int, Int)): DateTime = new DateTime(poiCell(address).getDateCellValue.getTime)

  private def poiCell(address: (Int, Int)): PoiCell = {
    val (c, r) = address
    sheet.getRow(r).getCell(c)
  }

  def exists(cell: String): Boolean = {
    exists(Cell.split(cell))
  }

  private def exists(address: (Int, Int)): Boolean = {
    val (c, r) = address
    sheet.getRow(r) match {
      case null => false
      case row => row.getCell(c) != null
    }
  }
}

object Cell {
  def split(cell: String): (Int, Int) = {
    val c = cell.filter(c => !c.isDigit)
    val r = cell.filter(c => c.isDigit)
    (toColumnIndex(c), r.toInt - 1)
  }

  private val columMap = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".zipWithIndex.map{ case (a, n) => (a, n)}.toMap

  private def toColumnIndex(column: String): Int = {
    column.reverse.map(c => columMap.get(c).get).zipWithIndex.map{ case (v, n) =>
      val f = (x: Int) => x * 26
      n match {
        case 0 => v
        case l => (1 to l).fold(v + 1)((y, r) => {
          f(y)
        })
      }
    }.sum
  }
}

object Book {
  def apply(path: String): Book = {
    new Book(poiWorkbook(path))
  }

  def apply(file: File): Book = {
    new Book(poiWorkbook(file))
  }

  def poiWorkbook(path: String): PoiWorkbook = {
    val stream = new File(path).exists() match {
      case true => new FileInputStream(path)
      case false => getClass().getResourceAsStream(path)
    }
    workbook(path, stream)
  }

  def poiWorkbook(file: File): PoiWorkbook = workbook(file.getAbsolutePath, new FileInputStream(file))

  private def workbook(path: String, stream: InputStream): PoiWorkbook = {
    path match {
      case p if p.endsWith(".xlsx") => new XSSFWorkbook(stream)
      case p if p.endsWith(".xls") => new HSSFWorkbook(stream)
    }
  }
}
