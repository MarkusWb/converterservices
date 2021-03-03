# converterservices

A set of wrappers (steps) for the conversion of commonly used data formats (csv, json, excel, xml)


The list of supported steps contains

* Saxon XSLT
* Saxon XQuery
* FOP XSL-FO
* PDF splitting and merging (pdfbox)
* XSLX/CSV <-> XML (Apache POI)
* Zipping/unzipping
* md5 hashing
* exiftool meta data extraction
* Upload and download by FTP and HTTP
* Send mails (Apache commons mail)

*converterservices* can be
* used as library
* started as web service, using [sparkjava](http://sparkjava.com/) as framework
* excecuted from the command line 

Error reporting can be configured (send the reports to a remote server logging instance or by email).

Check out the [Wiki](https://github.com/MarkusWb/converterservices/wiki) and the [JavaDoc](https://MarkusWb.github.io/converterservices/)
to find out more about *converterservices*.


*converterservices* requires Java 8.

Some wrapped external tools have to be installed independently and made accessible (ImageMagick, exiftool) if they shall be used,
but most of the functionality makes use of open source Java libraries which can be bundled in a single JAR (using the Gradle fatJar task).

In order to use the provided `build.gradle` add a `gradle.properties` file to the root directory, containing the following lines:

> mavenUrl=your.maven.repo

> mavenUser=yourUsername

> mavenPassword=dummy

> mavenKeyfile=dummy

**converterservices is a WORK IN PROGRESS! Functionality, interfaces, and function signatures are subject to change even in minor version steps.**

**NOTE:** As *converterservices* is a toolbox it heavily makes use of static functions.
Nevertheless, the classes have no private constructors because one of our scenarios of use is calling the functions in a BaseX XQuery
environment by Java Binding which demands public constructors.
