organization := "com.github.saisse"

name := "excel-scala"

version := "0.1"

scalaVersion := "2.13.5"

libraryDependencies ++= Seq(
   "org.apache.poi" % "poi" % "5.0.0"
 , "org.apache.poi" % "poi-ooxml" % "5.0.0"
 , "joda-time" % "joda-time" % "2.10.10"
 , "org.scalatest" %% "scalatest" % "3.2.7" % "test"
)

crossScalaVersions := Seq("2.10.0", "2.11.0", "2.12.0", "2.13.0")
