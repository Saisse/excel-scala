package com.github.saisse.excel_scala

import org.joda.time.DateTime
import org.joda.time.format.DateTimeFormat
import org.scalatest.{FlatSpec, FunSpec, Matchers}

class BookSpec extends FunSpec with Matchers {

  it("read xlsx file from file path") {
    val book = Book("./src/test/resources/test.xlsx")
  }

  it("read xlsx file with password") {
    val book = Book("./src/test/resources/password.xlsx", "test")
  }

  it("read xlsx file with invalid password") {
    intercept[Exception](Book("./src/test/resources/password.xlsx", "zzz"))
  }

  it("read xlsx file from classpath") {
    val book = Book("/test.xlsx")
  }

  it("read xls file") {
    val book = Book("/test.xls")
  }
}

class SheetSpec extends FlatSpec with Matchers {

  it should "data" in {

    val book = Book("./src/test/resources/test.xlsx")
    val sheet = book.sheet("Data")

    sheet.string("A2") should be ("A")
    sheet.string("A3") should be ("テスト")
    sheet.string("A4") should be ("Aテスト")
    sheet.string("A5") should be ("12345")
    sheet.string("A6") should be ("0100")
    sheet.string("A7") should be ("0100")
    sheet.string("A8") should be ("13802425287")

    sheet.stringOpt("A4") should be (Some("Aテスト"))
    sheet.stringOpt("A9") should be (None)

    sheet.double("B2") should be (2)
    sheet.double("B3") should be (3.4)
    sheet.double("B4") should be (5.4)

    sheet.doubleOpt("B4") should be (Some(5.4))
    sheet.doubleOpt("B5") should be (None)

    sheet.int("B2") should be (2)
    sheet.int("B3") should be (3)
    sheet.int("B4") should be (5)

    sheet.intOpt("B4") should be (Some(5))
    sheet.intOpt("B5") should be (None)

    sheet.dateTime("C2") should be(DateTime.parse("2014/12/30", DateTimeFormat.forPattern("yyyy/MM/dd")))
    sheet.dateTimeOpt("C2") should be(Some(DateTime.parse("2014/12/30", DateTimeFormat.forPattern("yyyy/MM/dd"))))
    sheet.dateTimeOpt("C3") should be(None)
  }

  it should "listRows" in {

    val book = Book("./src/test/resources/test.xlsx")
    val sheet = book.sheet("listRows")

    sheet.listRows(1, (row: Int) => sheet.exists(s"A$row")) should be (Seq(1, 2, 3))
    sheet.listRows(2, (row: Int) => sheet.exists(s"B$row")) should be (Seq(2, 3, 4))
  }

  it should "parse Row" in {

    val book = Book("./src/test/resources/test.xlsx")
    val sheet = book.sheet("parse")

    val result = sheet.parseRows(1, (row: Int) => sheet.exists(s"A$row"))((row: Int) => {
     Data(sheet.int(s"A$row"), sheet.string(s"B$row"))
    })

    result.length should be (2)
    result(0) should be (Data(1, "test"))
    result(1) should be (Data(2, "hoge"))

  }
  case class Data(i: Int, s: String)
}

class CellSpec extends FlatSpec with Matchers {
  it should "split" in {
    Cell.split("A1") should be ((0, 0))
    Cell.split("B1") should be ((1, 0))
    Cell.split("Z1") should be ((25, 0))
    Cell.split("AA1") should be ((26, 0))
    Cell.split("AB1") should be ((27, 0))
    Cell.split("BA1") should be ((52, 0))
    Cell.split("BB1") should be ((53, 0))
  }
}
