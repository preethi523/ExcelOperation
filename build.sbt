import sbt.errorssummary.Plugin.autoImport._

name := "exceloperations"

version := "0.1"

scalaVersion := "2.13.8"

libraryDependencies +="org.apache.poi" % "poi-ooxml" % "3.17"

libraryDependencies += "org.scalatest" %% "scalatest" % "3.2.9" % Test

libraryDependencies += "com.typesafe.scala-logging" %% "scala-logging" % "3.9.4"

libraryDependencies += "ch.qos.logback" % "logback-classic" % "1.2.10"

reporterConfig := reporterConfig.value.withColors(false)



reporterConfig := reporterConfig.value.withShortenPaths(false)



reporterConfig := reporterConfig.value.withColumnNumbers(false)



reporterConfig := reporterConfig.value.withReverseOrder(false)



reporterConfig := reporterConfig.value.withShowLegend(false)