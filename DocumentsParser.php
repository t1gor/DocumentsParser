<?php

	/**
	 * DOCx files parser
	 *
	 * Supported elements:
	 *  - paragraphs (w:p),
	 *  - images (pic:pic),
	 *  - links (w:hyperlink),
	 *  - tables (w:t),
	 *  - boorkmarks (w:bookmarkStart),
	 *  - lists.
	 *
	 * TODO:
	 *  - add links filter
	 *  - use absolutely equal images as one (remove extra images)
	 *  - add i18n
	 *
	 * Known issues:
	 *  - memory consuming
	 *  - for big files, executing more than 30 sec (default timout time)
	 *
	 * @author 	Igor Timoshenkov
	 * @started 20.03.2013
	 * @version 0.2
	 * @email 	igor.timoshenkov@gmail.com
	 *
	 * OpenXML Standards - http://www.schemacentral.com/sc/ooxml/ss.html
	 */

	class DocumentsParser {

		public 		$filesDestinationFolder 	= "";
		public 		$exclude 					= array();
		public 		$errors 					= array();
		public 		$styles 					= array();
		protected 	$previousListState			= "";

		// default parser settings
		private 	$defaultSettings = array(
			'filesDestinationFolder'	=> 'images',
			'exclude'					=> array(),
			'force'						=> true
		);

		// default element styles
		private 	$defaultStyles = array(
			'table'	=> array(
				'border'	=> '1',
			),
			'ul' => array(
				'style' 	=> 'margin-top: 5px; list-style: inside;',
			),
			'li' => array(),
			'p' => array(
				'style' => 'text-align: justify;',
			),
			'img' => array(
				'style' => 'clear: both; margin: 5px; float: left;',
			),
			'h2' => array(
				'style'	=> 'margin-top: 20px; margin-bottom: 5px; clear: both;',
			),
			'h3' => array(
				'style' => 'margin-top: 20px; margin-bottom: 5px; clear: both;',
			),
			'h4' => array(
				'style' => 'margin-top: 20px; margin-bottom: 5px; clear: both; font-size: 15px;',
			)
		);
 
		/**
		 * Init parser
		 *
		 * @param string $settings 	- parser settings
		 *    - 'imagesFolder'		- path to the folder for images save
		 *    - 'exclude'			- array of strings to exclude from the content
		 *    - 'force' 			- change memory limit and etc.
		 *
		 * @param array $styles 	- default html elements styles (see defaultStyles)
		 *    - 'element'	=> array(
		 * 			'attribute'	=> 'values'
		 * 		)
		 */
		public function __construct($settings = array(), $styles = array())
		{
			// merge settings
			$settings 	= array_merge($this->defaultSettings, $settings);
			$styles 	= array_merge($this->defaultStyles, $styles);

			if ($settings['force']) {
				error_reporting(E_ALL);
				set_time_limit(0);
				ini_set('memory_limit', '10000M');
				ini_set('max_execution_time', '0');
			}

			// check destination folder to be writable
			if (isset($settings['filesDestinationFolder']) && is_dir($settings['filesDestinationFolder']) && is_writable($settings['filesDestinationFolder'])) {
				$this->filesDestinationFolder = rtrim($settings['filesDestinationFolder'], '/')."/";
			}
			else {
				$this->errors[] = "The images folder should be writable. Please, check the permissions.";
				return;
			}

			$this->exclude 	= $settings['exclude'];
			$this->styles 	= $styles;
		}

		/**
		 * Parse single file
		 *
		 * @param string $filePath
		 *
		 * @return string
		 */
		public function parseFile($filePath)
		{
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

							$type = (string) $attributes['Type'][0];
							$tmp = explode("/", $type);
							$type = array_pop($tmp);

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

						$html = "";
						foreach ($docXML->xpath('w:body') as $parent)
						{
							$elementsCount = $this->countChildren($parent);

							foreach ($parent->xpath('*') as $sub)
							{
								switch ($sub->getName()) {

									case 'hyperlink':
										$html .= $this->parseHyperlink($sub);
									break;

									default:
									case 'r':
									case 'drawing':
										$html .= $this->parseImage($sub);
										$html .= $this->parseTextElement($sub);
									break;

									case 'p':
										$html .= $this->parsePragraph($sub);
									break;

									case 'tbl':
										if ($elementsCount > 50)  {
											$html .= $this->parseTable($sub);
										}
									break;
								}
								$this->filterContent($html);
							}
							if ($elementsCount < 50) {
								foreach ($docXML->xpath('w:body/w:tbl/w:tr') as $tblRow) {
									foreach ($tblRow->xpath('w:tc') as $tblCell) {
										foreach ($tblCell->xpath('*') as $cellContent)
										{
											switch ($cellContent->getName()) {

												case 'hyperlink':
													$html .= $this->parseHyperlink($cellContent);
												break;

												default:
												case 'r':
												case 'drawing':
													$html .= $this->parseImage($cellContent);
													$html .= $this->parseTextElement($cellContent);
												break;

												case 'p':
													$html .= $this->parsePragraph($cellContent);
												break;

												case 'tbl':
													$html .= $this->parseTable($cellContent);
												break;
											}
											$this->filterContent($html);
										}
									}
								}
							}
						}
					}
					// Close archive file
					$zip->close();

					// empty resources
					$this->resource_index = array();

					return $html;
				}
				else {
					$this->errors[] = "Couldn't open file {$filePath}.";
				}
			}
			return "";
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
		protected function parseTable($tableXML)
		{
			$content = "";

			// close the list if needed
			if ($this->previousListState) {
				$content .= "</ul>";
				$this->previousListState = false;
			}

			// start a table
			$content .= "<table ".$this->getAttributes('table').">";
			foreach ($tableXML->xpath('w:tr') as $tableRow) {
				$row = "<tr ".$this->getAttributes('td').">";
				$cellsPerRow = 0;
				foreach ($tableRow->xpath('w:tc') as $tableCell) {
					$cellsPerRow++;
					$merge = 0;
					$colspanFound = $tableCell->xpath('w:tcPr/w:gridSpan/@w:val');

					if (isset($colspanFound)) {
						foreach ($tableCell->xpath('w:tcPr/w:gridSpan/@w:val') as $colspan) {
							$merge = (string) $colspan[0];
						}
					}
					if ($merge > 0) {
						$row .= "<td colspan='{$merge}' ".$this->getAttributes('td').">";
						$tdTag = "<td colspan='{$merge}' ".$this->getAttributes('td').">";
					}
					else {
						$row .= "<td ".$this->getAttributes('td').">";
						$tdTag = "<td ".$this->getAttributes('td').">";
					}

					foreach ($tableCell->xpath('*') as $element)
					{
						switch ($element->getName()) {

							case 'tbl':
								$row .= $this->parseTable($element);
							break;

							case 'r':
								$row .= $this->parseImage($element, true);
								$row .= $this->parseTextElement($element);
							break;

							default:
								$row .= $this->parsePragraph($element, true);
							break;
						}
					}
					// remove first <br> in a cell
					$row = preg_replace("@".$tdTag."<br/>@", $tdTag, $row);
					$row .= "</td>";
				}
				$row .= "</tr>";

				$stripped = strip_tags($row);
				$hasImages = strpos($row, '<img');

				// center single image in cell
				if (empty($stripped) && $cellsPerRow == 1 && $hasImages) {
					$row = preg_replace("@<br/>@", "", $row);
					$content .= str_replace("<p", "<p style='text-align: center;'", $row);
				}
				// remove empty rows
				elseif (!empty($stripped) || $hasImages) {
					$content .= $row;
				}
			}
			$content .= "</table><br/>";
			return $content;
		}

		/**
		 * Get attributes string for the specified element
		 *
		 * @param string $elementName
		 *
		 * @return string
		 */
		private function getAttributes($elementName)
		{
			$attrString = "";
			foreach ($this->styles[$elementName] as $key => $value) {
				$attrString .= $key ."=\"{$value}\"";
			}
			return $attrString;
		}

		/**
		 * If an html part starts with the br tag - remove it
		 *
		 * @param string $html
		 *
		 * @return void
		 */
		protected function removeFirstBr($html)
		{
			$firstBrPos = strpos($html, "<br/>");
			if ($firstBrPos == 0 && $firstBrPos !== false) {
				$html = substr($html, 4, strlen($html));
			}
			return $html;
		}

		/**
		 * Parse link element
		 *
		 * @param SimpleXMLElement $xmlElement
		 *
		 * @return string (html)
		 */
		protected function parseHyperlink($xmlElement)
		{
			$text = $xmlElement->xpath('w:r/w:t');
			if (isset($text) && gettype($text) != 'NULL') {

				$text_string = "";
				foreach ($text as $char) {
					$text_string .= $char;
				}
				$text_string = trim($text_string);

				// remove wiki brackets
				$text_string = str_replace("[", "", $text_string);
				$text_string = str_replace("]", "", $text_string);

				$href = $this->resource_index[(string) $xmlElement->xpath('@r:id')[0]['id']]['target'];
				return "<a href='".$href."' target='_blank' ".$this->getAttributes('a').">".$text_string."</a>";
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
		protected function parsePragraph($xmlElement, $inline = false)
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

						case 'bookmarkStart':
							$p_html .= $this->parseBookmark($sub_element);
						break;

						case 'tbl':
							$p_html .= $this->parseTable($element);
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

			if (!in_array(trim(strip_tags($p_html)), $this->exclude)) {
				// check for size
				$xmlString = $xmlElement->saveXML();
				$h2BySize 	= (bool) strpos($xmlString, "w:val=\"Heading2\"");
				$h3BySize 	= (bool) strpos($xmlString, "w:val=\"Heading3\"");
				$h4BySize 	= (bool) strpos($xmlString, "w:val=\"Heading4\"");
				$bold 		= (bool) strpos($xmlString, "w:val=\"Strong\"");
				$italic 	= (bool) strpos($xmlString, "w:val=\"Emphasis\"");
				$underlined = (bool) strpos($xmlString, "<w:u ");

				$prefix = "";

				// handle list if needed
				if ($isListElement && !empty($p_html)) {
					// previous paragraph is also a list element?
					if ($this->previousListState) {
						return "<li ".$this->getAttributes('li').">".$p_html."</li>";
					}
					else {
						$this->previousListState = true;
						return "<ul ".$this->getAttributes('ul')."><li ".$this->getAttributes('li').">".$p_html."</li>";
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
					if ($h2BySize || $h3BySize || $h4BySize || $bold || $italic || $underlined)
					{
						$tag = null;
						$style = "";

						if ($italic) {
							$tag = "i";
						}
						if ($italic) {
							$tag = "u";
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
							return $prefix."<{$tag} ".$this->getAttributes($tag).">".$p_html."</{$tag}>";
						}
						else {
							return $prefix."<p ".$this->getAttributes('p').">".$p_html."</p>";
						}
					}
					else {
						// try to get text size
						if ($xmlElement->xpath("w:pPr/w:rPr/w:sz/@w:val")[0]) {
							$sizeStr = $xmlElement->xpath("w:pPr/w:rPr/w:sz/@w:val")[0]->saveXML();
							$sizeStrParts = explode("=\"", $sizeStr);
							$size = (int) trim($sizeStrParts[1], "\"") / 2;
							if ($size > 12 && $size < 18) {
								return $prefix."<h4 ".$this->getAttributes('h4').">".$p_html."</h4>";
							}
						}
						return $prefix."<p ".$this->getAttributes('p').">".$p_html."</p>";
					}
				}
				else {
					return $prefix."<br/>";
				}
			}
			return $prefix."";
		}

		/**
		 * Remove symbols like Â, etc.
		 *
		 * @param string $string
		 *
		 * @return string
		 */
		private function normalize($string) {
			return mb_convert_encoding($string, "HTML-ENTITIES", "UTF-8");
		}

		/**
		 * Simply get text out of an element
		 *
		 * @param SimpleXmlElement $xmlElement
		 *
		 * @return string
		 */
		protected function parseTextElement($xmlElement)
		{
			$text = "";

			$italic 	= (strpos($xmlElement->saveXML(), "<w:i/>") !== false);
			$bold 		= (strpos($xmlElement->saveXML(), "<w:b/>") !== false);
			$underline 	= (strpos($xmlElement->saveXML(), "<w:u/>") !== false);

			foreach ($xmlElement->xpath('w:t') as $textElement) {
				$text .= strip_tags($textElement->saveXML());
			}

			// exclude if needed
			if (in_array($text, $this->exclude)) {
				return "";
			}

			// apply styles
			$text = $this->normalize($text);
			$text = $italic 	? "<i ".$this->getAttributes('i').">".$text."</i>" : $text;
			$text = $bold 		? "<b ".$this->getAttributes('b').">".$text."</b>" : $text;
			$text = $underline 	? "<u ".$this->getAttributes('u').">".$text."</u>" : $text;

			return $text;
		}

		/**
		 * Parse images
		 *
		 * @param SimpleXMLElement $xmlElement
		 * @param bool 			   $inline
		 *
		 * @return string (html)
		 */
		protected function parseImage($xmlElement, $inline = false)
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
						$src = $this->resource_index[$id]['target'];
						$src = $this->filesDestinationFolder.str_replace('media/', '', $src);
					}
					return "<img src='{$src}' ".$this->getAttributes('img')." />";
				}
			}

			// Epic images in shapes part goes here:

			foreach ($xmlElement->xpath('w:pict') as $picture) {
				// iterate through groups
				foreach ($picture->xpath('v:group') as $group) {
					// ... and sub groups
					foreach ($group->xpath('v:group') as $subGroup) {
						foreach ($subGroup->xpath('v:shape') as $shape) {
							foreach ($subGroup->xpath('*/v:imagedata') as $imagedata) {
								$imageId = (string) $imagedata->xpath('@r:id')[0];
								$src = $this->resource_index[$imageId]['target'];
								$src = $this->filesDestinationFolder.str_replace('media/', '', $src);
								return "<img src='{$src}' ".$this->getAttributes('img')." />";
							}
						}
					}

					// but don't also forget about images
					foreach ($group->xpath('v:shape') as $shape) {
						foreach ($group->xpath('*/v:imagedata') as $imagedata) {
							$imageId = (string) $imagedata->xpath('@r:id')[0];
							$src = $this->resource_index[$imageId]['target'];
							$src = $this->filesDestinationFolder.str_replace('media/', '', $src);
							return "<img src='{$src}' ".$this->getAttributes('img')." />";
						}
					}
				}

				// but don't also forget about images
				foreach ($picture->xpath('v:shape') as $shape) {
					foreach ($picture->xpath('*/v:imagedata') as $imagedata) {
						$imageId = (string) $imagedata->xpath('@r:id')[0];
						$src = $this->resource_index[$imageId]['target'];
						$src = $this->filesDestinationFolder.str_replace('media/', '', $src);
						return "<img src='{$src}' ".$this->getAttributes('img')." />";
					}
				}
			}

			return "";
		}

		/**
		 * Count Simple Xml Element children
		 *
		 * @param SimpleXmlElement $parent
		 *
		 * @return int
		 */
		private function countChildren($parent)
		{
			$elementsCount = 0;
			foreach ($parent->xpath('*') as $sub) {
				$elementsCount++;
			}
			return $elementsCount;
		}

		/**
		 * Handle multiple linebreaks, remove spear ones
		 *
		 * @param string $html
		 *
		 * @return void
		 */
		protected function filterContent(&$html)
		{
			// many <br>'s to one
			$html = preg_replace("@(<br/>)+@", "<br/>", $html);

			// remove the first br
			if (strpos($html, "<br/>") !== false && strpos($html, "<br/>") == 0) {
				$html = substr($html, 5);
			}

			// remove extra <br> before tables
			$html = preg_replace("@(<br/><table)+@", "<table", $html);

			// remove extra <br> before and after headers - margins are set there
			$html = preg_replace("@(<br/><h1)+@", "<h1", $html);
			$html = preg_replace("@(<br/><h2)+@", "<h2", $html);
			$html = preg_replace("@(<br/><h3)+@", "<h3", $html);
			$html = preg_replace("@(<br/><h4)+@", "<h4", $html);
			$html = preg_replace("@(</h1><br/>)+@", "</h1>", $html);
			$html = preg_replace("@(</h2><br/>)+@", "</h2>", $html);
			$html = preg_replace("@(</h3><br/>)+@", "</h3>", $html);
			$html = preg_replace("@(</h4><br/>)+@", "</h4>", $html);
		}

	}

?>
