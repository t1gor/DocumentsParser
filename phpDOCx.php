<?php

    /**
     * DOCx files parser
	 *
	 * Supported elements: paragraphs (w:p), images (pic:pic), links (w:hyperlink), tables (w:t), boorkmarks (w:bookmarkStart), lists.
	 *
	 * @author 	Igor Timoshenkov
	 * @started 20.03.2013
	 * @email 	igor.timoshenkov@gmail.com
	 *
	 * OpenXML Standards - http://www.schemacentral.com/sc/ooxml/ss.html
	 */

	class DocumentsParser {

		public $filesDestinationFolder 	= "";
		protected $parseLastError 		= "";
		protected $previousListState	= "";

		public function __construct($imagesFolder)
		{
			error_reporting(E_ALL);
			set_time_limit(0);
			ini_set('memory_limit', '20000M');
			ini_set('max_execution_time', '0');
			$this->filesDestinationFolder = $imagesFolder;
		}

		/**
		 * Parse single file
		 *
		 * @param string $filePath
		 *
		 * @return array
		 */
		function parseFile($filePath)
		{
			$mime = mime_content_type($filePath);
			if (is_file($filePath)) {

				$zip = new ZipArchive;
    			$dataFile 		= 'word/document.xml';
    			$resourseFile 	= 'word/_rels/document.xml.rels';

    			if (true === $zip->open($filePath)) {

    				// search for resourses file
    				$this->resource_index = array();
    				if (($resourses = $zip->locateName($resourseFile)) !== false)
					{
						$resouse = $zip->getFromIndex($resourses);
						$resXML = simplexml_load_string($resouse, 'SimpleXMLElement');
						$this->setNameSpaces($resXML);

						foreach ($resXML->children() as $rel) {
							$attributes = $rel->attributes();

							$type 	= (string) $attributes['Type'][0];
							$tmp 	= explode("/", $type);
							$type 	= array_pop($tmp);

							if (in_array($type, array('image', 'hyperlink'))) {

								$target = (string) $attributes['Target'][0];

								$this->resource_index[(string) $attributes['Id'][0]] = array(
									'type'		=> $type,
									'target'	=> $target,
									'external'	=> isset($attributes['TargetMode']) && $attributes['TargetMode'] == 'External'
								);

								// copy files if image
								if ('image' == $type) {
									$fp = $zip->getStream('word/'.$target);

									ob_start();
										fpassthru($fp);
										$image = ob_get_contents();
									ob_end_clean();

									// create dir if needed
									if (!is_dir($this->filesDestinationFolder)) {
										mkdir($this->filesDestinationFolder);
									}
									file_put_contents($this->filesDestinationFolder.str_replace('media/', '', $target), $image);
								}
							}
						}
					}

					// If done, search for the data file in the archive
					if (($index = $zip->locateName($dataFile)) !== false)
					{
						// If found, read it to the string
						$data = $zip->getFromIndex($index);

						// Load XML from a string
						$docXML = simplexml_load_string($data, 'SimpleXMLElement');
						$this->setNameSpaces($docXML);

						// iterate through DOM elements
						$content = "";
						foreach ($docXML->xpath('w:body') as $parent)
						{
							foreach ($parent->xpath('*') as $sub)
							{
								switch ($sub->getName()) {

									case 'hyperlink':
										$content .= $this->parseHyperlink($sub);
									break;

									default:
									case 'r':
									case 'drawing':
										$content .= $this->parseImage($sub);
										$content .= $this->parseTextElement($sub);
									break;

									case 'p':
										$content .= $this->parsePragraph($sub);
									break;

									case 'tbl':
										$content .= $this->parseTable($sub);
									break;
								}
							}
						}
					}
					// Close archive file
					$zip->close();

					// empty resources
					$this->resource_index = array();

					return $content;
				}
				else {
					error_log("Couldn't open file {$filePath} ({$mime})");
				}
			}
			return array();
		}

		/**
		 * Remove symbols like Â, etc.
		 *
		 * @param string $string
		 *
		 * @return string
		 */
		function normalize($string) {
			return mb_convert_encoding($string, "HTML-ENTITIES", "UTF-8");
		}

		/**
		 * Add namespaces to SimpleXmlElement
		 *
		 * @param SimpleXmlElement $xml
		 *
		 * @return false
		 */
		protected function setNameSpaces($xml)
		{
			$xml->registerXPathNamespace("a", 	"http://schemas.openxmlformats.org/drawingml/2006/main");
			$xml->registerXPathNamespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
			$xml->registerXPathNamespace("r", 	"http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			$xml->registerXPathNamespace("w", 	"http://schemas.microsoft.com/office/word/2003/wordml");
			$xml->registerXPathNamespace("wp", 	"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
		}

		/**
		 * Parse table element
		 *
		 * @param SimpleXMLElement $tableXML
		 *
		 * @return string
		 */
		function parseTable($tableXML)
		{
			// close the list if needed
			if ($this->previousListState) {
				$content .= "</ul>";
				$this->previousListState = false;
			}

			// start a table
			$content .= "<table border='1'>";
			foreach ($tableXML->xpath('w:tr') as $tableRow) {
				$content .= "<tr>";
				foreach ($tableRow->xpath('w:tc') as $tableCell) {
					$content .= "<td>";
					foreach ($tableCell->xpath('*') as $element) {
						$content .= $this->parsePragraph($element, true);
					}
					$content .= "</td>";
				}
				$content .= "</tr>";
			}
			$content .= "</table>";
			return $content;
		}

		/**
		 * Parse link element
		 *
		 * @param SimpleXMLElement $xmlElement
		 *
		 * @return string (html)
		 */
		function parseHyperlink($xmlElement)
		{
			$text = $xmlElement->xpath('w:r/w:t');
			if (isset($text) && gettype($text) != 'NULL') {

				$text_string = "";
				foreach ($text as $char) {
					$text_string .= $char;
				}

				$href = $this->resource_index[(string) $xmlElement->xpath('@r:id')[0]['id']]['target'];
				return "<a href='".$href."' target='_blank'>".$text_string."</a>";
			}
			return "";
		}

		/**
		 * Parse paragraph element
		 *
		 * @param SimpleXMLElement $xmlElement
		 * @param bool 			   $inline
		 *
		 * @return void
		 */
		function parsePragraph($xmlElement, $inline = false)
		{
			$p_html = "";
			$children = $xmlElement->xpath('*');

			// check if it's a list item
			$isListElement = (bool) strpos($xmlElement->saveXML(), "<w:numPr>");

			if (!empty($children)) {
				foreach ($children as $sub_element)
				{
					switch ($sub_element->getName()) {

						case 'hyperlink':
							$p_html .= $this->parseHyperlink($sub_element);
						break;

						default:
						case 'r':
							$p_html .= $this->parseImage($sub_element, ($inline || $isListElement));
							$p_html .= $this->parseTextElement($sub_element);
						break;

						case 'drawing':
							$p_html .= $this->parseImage($sub_element);
						break;

						case 't':
							$p_html .= $this->parseTextElement($sub_element);
						break;
					}
				}
			}

			$p_html = $this->normalize($p_html);

			// check for header style
			$isHeader = (bool) strpos($xmlElement->saveXML(), "mw-headline");

			if (!in_array($p_html, $this->exclude)) {
				// check for size (header size = 36)
				if (isset($xmlElement->xpath('w:r')[0])) {
					$xmlString = $xmlElement->xpath('w:r')[0]->saveXML();
					$h2BySize 	= strpos($xmlString, "w:val=\"36\"");
					$h3BySize 	= strpos($xmlString, "w:val=\"30\"");
					$h4BySize 	= strpos($xmlString, "w:val=\"26\"");
					$bBySize 	= strpos($xmlString, "<w:b/>");
				}
				$bookmark = str_replace(" ", "_", $p_html);
				$prefix = "";

				// handle list if needed
				if ($isListElement) {
					// previous paragraph is also a list element?
					if ($this->previousListState) {
						return "<li>".$p_html."</li>";
					}
					else {
						$this->previousListState = true;
						return "<ul><li>".$p_html."</li>";
					}
				}
				else {
					// was previous a list element?
					if ($this->previousListState) {
						$prefix = "</ul>";
					}
				}
				$this->previousListState = $isListElement;

				if (!empty($p_html)) {
					if (isset($xmlElement->xpath('w:r')[0]) && ($h2BySize || $h3BySize || $h4BySize)) {
						$tag = null;
						if ($bBySize) {
							$tag = "b";
						}
						if ($h2BySize) {
							$tag = "h2";
						}
						if ($h3BySize) {
							$tag = "h3";
						}
						if ($h4BySize) {
							$tag = "h4";
						}
						if (!is_null($tag)) {
							return $prefix."<{$tag}>".$p_html."</{$tag}>";
						}
						else {
							return $prefix."<p>".$p_html."</p>";
						}
					}
					else {
						return $prefix."<p>".$p_html."</p>";
					}
				}
			}
			return $prefix."";
		}

		/**
		 * Simply get text out of an element
		 *
		 * @param SimpleXmlElement $xmlElement
		 *
		 * @return string
		 */
		function parseTextElement($xmlElement)
		{
			$text = "";
			foreach ($xmlElement->xpath('w:t') as $textElement) {
				$text .= strip_tags($textElement->saveXML());
			}
			return $this->normalize($text);
		}

		/**
		 * Parse images
		 *
		 * @param SimpleXMLElement $xmlElement
		 * @param bool 			   $inline
		 *
		 * @return string (html)
		 */
		function parseImage($xmlElement, $inline = false)
		{
			preg_match("@<a:blip .*>@", $xmlElement->saveXML(), $matches);

			if (!empty($matches)) {
				$parts = explode("/>", $matches[0]);
				$blipXML = simplexml_load_string($parts[0]."/>");
				if (gettype($blipXML) == 'object') {
					$this->setNameSpaces($blipXML);

					$id = (string) $blipXML->attributes()['r:embed'];

					if ($this->resource_index[$id]['external']) {
						$src = $this->resource_index[$id]['target'];
					}
					else {
						$src = $this->filesDestinationFolder.str_replace('media/', '', $this->resource_index[$id]['target']);
					}
					$float = !$inline ? 'margin: 5px; float: left;' : '';
					return "<img src='{$src}' style='{$float}' />";
				}
			}
			return "";
		}
	}
?>
