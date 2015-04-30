lazy val root = (project in file(".")).
  settings(
    name:="excelToXml",
    version := "1.0",
    crossPaths := false,
    autoScalaLibrary := false,
    mainClass in (Compile, run) := Some("edu.ur.dh.xmler.Xmler"),
    libraryDependencies ++= Seq(
      "org.apache.poi" % "poi" % "3.11",
      "commons-codec" % "commons-codec" % "1.9",
      "org.apache.poi" % "poi-ooxml" % "3.11",
      "org.apache.poi" % "poi-ooxml-schemas" % "3.11",
      "org.apache.xmlbeans" % "xmlbeans" % "2.6.0",
      "stax" % "stax-api" % "1.0.1",
      "commons-cli" % "commons-cli" % "1.2",
      "commons-io" % "commons-io" % "2.4"
    )
    
  )
  

javacOptions ++= Seq("-source", "1.6", "-target", "1.6")
