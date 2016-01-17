package com.debasishalder.utility{
	import flash.system.System;
	import flash.filesystem.File;
	import flash.filesystem.FileStream;
	import flash.filesystem.FileMode;

	public class speradsheetParser {
		private var XMLForUpdate: XML
		public function speradsheetParser(Strng: String) {
			var str: String = Strng;
			str = str.split("'").join("&#39;");
			str = str.split('x:').join('');
			str = str.split('o:').join('');
			str = str.split('ss:').join('');
			str = str.split('html:').join('');
			str = str.substring(str.indexOf('<Worksheet'), str.indexOf('</Worksheet>') + 12);
			str = str.substring(str.indexOf('<Row'), str.lastIndexOf('</Row>') + 6);
			str = str.split('xmlns="urn:schemas-microsoft-com:office:spreadsheet"').join('');
			str = str.split('xmlns:o="urn:schemas-microsoft-com:office:office"').join('');
			str = str.split('xmlns:x="urn:schemas-microsoft-com:office:excel"').join('');
			str = str.split('xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"').join('');
			str = str.split('xmlns:html="http://www.w3.org/TR/REC-html40"').join('');
			str = str.split('xmlns="http://www.w3.org/TR/REC-html40"').join('');
			str = str.split('&#10;').join('<br/>');
			var xx: XML = new XML('<db>' + str + '</db>');
			str = xx.toString();
			System.disposeXML(xx);
			xx = null;
			str = str.replace(/(>)\s*\n\s*(<)/ig, '>  <');
			str = str.replace(/(>)\n(<)/ig, '><');
			str = str.replace(/(>)\n\s*(<)/ig, '><');
			str = str.replace(/(>)\s*\n(<)/ig, '><');
			str = str.replace(/<\s*?Data[^<>]*>/ig, '<Data><![CDATA[');
			str = str.replace(/<\s*?\/\Data[^<>]*>/ig, ']]></Data>');
			try {
				XMLForUpdate = new XML(str);
			} catch (e: Error) {
				throw new Error('Got an error on XML processing Please check the XML')
			}
		}
		
		public static function fileToString(FILES:File, FILESTREAMS:FileStream = null): String {
			var isStreamCreated: Boolean = false;
			if (FILESTREAMS == null) {
				isStreamCreated = true;
				FILESTREAMS = new FileStream();
			}
			FILESTREAMS.open(FILES, FileMode.READ);
			var fileContents: String = FILESTREAMS.readUTFBytes(FILESTREAMS.bytesAvailable);
			FILESTREAMS.close();
			if (isStreamCreated) {
				FILESTREAMS = null;
			}
			return fileContents;
		}

		public function getRow(RowNum: int): Array {

			var RemoveFonttag: Boolean = true
			var XMLRowArray: Array = new Array();
			if (RowNum >= XMLForUpdate.children().length()) {
				XMLRowArray.push('')
				return XMLRowArray
			}
			var NumColumn: int = XMLForUpdate.children()[RowNum].children().length();
			for (var j: int = 0; j < NumColumn; j++) {
				var CellString: String = XMLForUpdate.children()[RowNum].children()[j].children().children();
				if (RemoveFonttag) {
					var fonttag: RegExp = new RegExp('\<(FONT|font)[^>]*\>', "gi");
					var fonttagcls: RegExp = new RegExp("\(</FONT>|</font>)", "gi");
					CellString = CellString.replace(fonttag, '');
					CellString = CellString.replace(fonttagcls, '');
				}

				CellString = CellString.replace(/\s*<Sub>/ig, '<Sub>');
				CellString = CellString.replace(/\s*<Sup>/ig, '<Sup>');
				var findNewLineRegExp: RegExp = new RegExp("\n", "g");
				CellString = CellString.replace(findNewLineRegExp, '');
				CellString = CellString.split('  ').join(' ');
				var chkstr: String = ''
				if (CellString != null) {
					chkstr = CellString.split(' ').join('')
					var smlchkstr: String = chkstr.toLowerCase()
				}
				if (CellString == '' || CellString == null || chkstr == '') {
					CellString = 'BLANK';
					if (chkstr == 'blank' || smlchkstr == 'null' || smlchkstr == 'undefined') {
						CellString = 'null'
					}
				}
				var IndexVal: String = String(XMLForUpdate.children()[RowNum].children()[j].attribute('Index'));
				if (IndexVal == '' || IndexVal == null) {
					CellString = CellString.split("'").join('&#39;');
					CellString = CellString.split("&amp;").join('&');
					XMLRowArray.push(CellString);
				} else {
					if ((XMLRowArray.length) != Number(IndexVal)) {
						for (var k: int = XMLRowArray.length + 1; k < Number(IndexVal); k++) {
							XMLRowArray.push('BLANK');
						}
						CellString = CellString.split('&amp;').join('&');
						XMLRowArray.push(CellString);
					}
				}
			}
			
			return XMLRowArray;
		}
		
		public function get rowLength(): int {
			var N: int = XMLForUpdate.children().length();
			return N;
		}

		
		

}

}