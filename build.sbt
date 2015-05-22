organization := "com.github.saisse"

name := "excel-scala"

version := "0.1"

scalaVersion := "2.11.6"


libraryDependencies ++= Seq(
    "org.apache.poi" % "poi" % "3.11"
  , "org.apache.poi" % "poi-ooxml" % "3.11"
  , "joda-time" % "joda-time" % "2.5"
  , "org.scalatest" %% "scalatest" % "2.2.4" % "test"
)

crossScalaVersions := Seq("2.10.0", "2.11.0")
