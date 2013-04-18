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