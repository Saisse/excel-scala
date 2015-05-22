package com.github.saisse.excel_scala

import java.io.{File, FileInputStream}
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.{Sheet => PoiSheet, Workbook => PoiWorkbook, Cell => PoiCell}
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

  def string(cell: String): String = {
    val address = Cell.split(cell)
    poiCell(address).getRichStringCellValue.getString
  }

  def double(cell: String): Double = {
    val address = Cell.split(cell)
    poiCell(address).getNumericCellValue
  }

  def int(cell: String): Int = {
    val address = Cell.split(cell)
    poiCell(address).getNumericCellValue.toInt
  }

  def dateTIme(cell: String): DateTime = {
    val address = Cell.split(cell)
    new DateTime(poiCell(address).getDateCellValue.getTime)
  }

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
    val stream = new File(path).exists() match {
      case true => new FileInputStream(path)
      case false => getClass().getResourceAsStream(path)
    }
    new Book(path match {
      case p if p.endsWith(".xlsx") => new XSSFWorkbook(stream)
      case p if p.endsWith(".xls") => new HSSFWorkbook(stream)
    })
  }
}
