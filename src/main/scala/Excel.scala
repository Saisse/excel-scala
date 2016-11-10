package com.github.saisse.excel_scala

import java.io.{File, FileInputStream, InputStream}

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.poifs.crypt.{Decryptor, EncryptionInfo}
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.apache.poi.ss.usermodel.{DataFormatter, Cell => PoiCell, Sheet => PoiSheet, Workbook => PoiWorkbook}
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.joda.time.DateTime

class Book(book: PoiWorkbook) {
  def sheet(name: String): Sheet = new Sheet(book.getSheet(name))
  def sheet(index: Int): Sheet = new Sheet(book.getSheetAt(index))
}

class Sheet(sheet: PoiSheet) {
  def listRows(startRow: Int, end: (Int) => Boolean): Seq[Int] = {
    end(startRow) match {
      case false => Nil
      case true => startRow +: listRows(startRow + 1, end)
    }
  }

  def poiSheet = sheet

  def parseRow[A](row: Int)(parser: (Int) => A): A = parser(row)
  def parseRows[A](startRow: Int, end: (Int) => Boolean)(parser: (Int) => A): Seq[A] = {
    listRows(startRow, end).map(row => parser(row))
  }

  def cell(column: String, row: Int): String = Cell.cell(column, row)

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
        val df = new DataFormatter()
        df.formatCellValue(c)
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

  def rowParser(row: Int) = new RowParser(this, row)

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

  def cell(column: String, row: Int): String = s"$column$row"

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

  def apply(path: String, password: String): Book = {
    new Book(poiWorkbook(path, Some(password)))
  }

  def apply(file: File): Book = {
    new Book(poiWorkbook(file))
  }

  def apply(fileName: String, file: File): Book = {
    new Book(poiWorkbook(fileName, file))
  }

  def apply(fileName: String, file: File, password: String): Book = {
    new Book(poiWorkbook(fileName, file, password))
  }

  def apply(file: File, password: String): Book = {
    new Book(poiWorkbook(file, password))
  }

  def poiWorkbook(path: String, password: Option[String] = None): PoiWorkbook = {
    val stream = new File(path).exists() match {
      case true => new FileInputStream(path)
      case false => getClass().getResourceAsStream(path)
    }
    workbook(path, stream, password)
  }

  def poiWorkbook(file: File): PoiWorkbook = poiWorkbook(file.getAbsolutePath, file)
  def poiWorkbook(file: File, password: String): PoiWorkbook = poiWorkbook(file.getAbsolutePath, file, password)
  def poiWorkbook(fileName: String, file: File): PoiWorkbook = workbook(fileName, new FileInputStream(file), None)
  def poiWorkbook(fileName: String, file: File, password: String): PoiWorkbook = workbook(fileName, new FileInputStream(file), Some(password))

  private def workbook(path: String, stream: InputStream, password: Option[String]): PoiWorkbook = {

    val s = password.map(p => withPassword(stream, p)).getOrElse(stream)
    path match {
      case p if p.endsWith(".xlsx") => new XSSFWorkbook(s)
      case p if p.endsWith(".xls") => new HSSFWorkbook(s)
    }
  }

  private def withPassword(stream: InputStream, password: String): InputStream = {
    val s = new POIFSFileSystem(stream)
    val enc = new EncryptionInfo(s)
    val decryptor = Decryptor.getInstance(enc)

    if(!decryptor.verifyPassword(password)) {
      throw new Exception("invalid password")

    }
    decryptor.getDataStream(s)
  }
}

class RowParser(sheet: Sheet, row: Int) {
  private def cell(column: String, row: Int) = Cell.cell(column, row)

  def string(column: String) = sheet.string(cell(column, row))
  def stringOpt(column: String) = sheet.stringOpt(cell(column, row))

  def double(column: String) = sheet.double(cell(column, row))
  def doubleOpt(column: String) = sheet.doubleOpt(cell(column, row))

  def int(column: String) = sheet.int(cell(column, row))
  def intOpt(column: String) = sheet.intOpt(cell(column, row))

  def dateTime(column: String) = sheet.dateTime(cell(column, row))
  def dateTimeOpt(column: String) = sheet.dateTimeOpt(cell(column, row))

  def exists(column: String) = sheet.exists(cell(column, row))
  def parse[A](row: Int)(parser: (Int) => A) = sheet.parseRow(row)(parser)
}
