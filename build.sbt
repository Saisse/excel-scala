organization := "com.github.saisse"

name := "excel-scala"

version := "0.1"

val scala212 = "2.12.10"
val scala213 = "2.13.6"
scalaVersion := scala213

libraryDependencies ++= Seq(
   "org.apache.poi" % "poi" % "5.0.0"
 , "org.apache.poi" % "poi-ooxml" % "5.0.0"
 , "joda-time" % "joda-time" % "2.10.10"
 , "org.scalatest" %% "scalatest" % "3.2.7" % "test"
)

releaseCrossBuild := true

crossScalaVersions := Seq(scala212, scala213)

publish / skip := true