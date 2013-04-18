DOCx files parser
========

Parses the docx file and returns an html string. Any help or critics appreciated.

Reminder links:
--------------------
* Comparing images - http://pastebin.com/wgeu2DqE
* OpenXML Standards - http://www.schemacentral.com/sc/ooxml/ss.html

Supported elements:
--------------------
* paragraphs (w:p),
* images (pic:pic),
* links (w:hyperlink),
* tables (w:t),
* boorkmarks (w:bookmarkStart),
* lists.

TODO:
--------------------
* add links filter
* images optimization (remove extra equal images)
* add i18n

Known issues:
--------------------
* memory consuming
* for big files, executing more than 30 sec (default timout time)

Usage example:
--------------------

```php
<?php
    // load lib
	require_once('DocumentsParser.php');

	// init parser
	$parserSettings = array(
		'filesDestinationFolder' => 'images',
	);

	$defaultStyles = array();

	$parser = new DocumentsParser($parserSettings, $defaultStyles);

	// parse DOCx
	$html = $parser->parseFile('test_document.docx');

	// save content to file
	file_put_contents('test_document.html', $html);
?>
```
