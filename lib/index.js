/*jshint globalstrict:true, devel:true */
/*global require, module, exports, process, __dirname, Buffer */
"use strict";

var fs = require( 'fs' ),
	path = require( 'path' ),
	zip = require( 'node-zip' ),
	etree = require( 'elementtree' );

module.exports = ( function () {

	var DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
		SHARED_STRINGS_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

	var Workbook = function ( data ) {
		var self = this;

		self.archive = null;
		self.sharedStrings = [];
		self.sharedStringsLookup = {};

		if ( data ) {
			self.loadTemplate( data );
		}
	};
	/* LOAD */
	Workbook.prototype.loadTemplate = function ( data ) {
		var self = this;

		if ( Buffer.isBuffer( data ) ) {
			data = data.toString( 'binary' );
		}

		self.archive = new zip( data, {
			base64: false,
			checkCRC32: true
		} );


		// Load relationships
		var rels = etree.parse( self.archive.file( "_rels/.rels" ).asText() ).getroot(),
			workbookPath = rels.find( "Relationship[@Type='" + DOCUMENT_RELATIONSHIP + "']" ).attrib.Target; //Gives something like : xl/workbook.xml


		self.workbookPath = workbookPath;
		self.prefix = path.dirname( workbookPath );
		self.workbook = etree.parse( self.archive.file( workbookPath ).asText() ).getroot();
		self.workbookRels = etree.parse( self.archive.file( self.prefix + "/" + '_rels' + "/" + path.basename( workbookPath ) + '.rels' ).asText() ).getroot();
		self.sheets = self.loadSheets( self.prefix, self.workbook, self.workbookRels );


		self.sharedStringsPath = self.prefix + "/" + self.workbookRels.find( "Relationship[@Type='" + SHARED_STRINGS_RELATIONSHIP + "']" ).attrib.Target;
		self.sharedStrings = [];
		etree.parse( self.archive.file( self.sharedStringsPath ).asText() ).getroot().findall( 'si/t' ).forEach( function ( t ) {
			self.sharedStrings.push( t.text );
			self.sharedStringsLookup[ t.text ] = self.sharedStrings.length - 1;
		} );
	};

	// Get a list of sheet ids, names and filenames
	Workbook.prototype.loadSheets = function ( prefix, workbook, workbookRels ) {
		var self = this;

		var sheets = [];

		workbook.findall( "sheets/sheet" ).forEach( function ( sheet ) {
			var sheetId = sheet.attrib.sheetId,
				relId = sheet.attrib[ 'r:id' ],
				relationship = workbookRels.find( "Relationship[@Id='" + relId + "']" ),
				filename = prefix + "/" + relationship.attrib.Target;

			sheets.push( {
				id: parseInt( sheetId, 10 ),
				name: sheet.attrib.name,
				filename: filename
			} );
		} );

		return sheets;
	};

	Workbook.prototype.loadSheet = function ( sheet ) {
		var self = this;

		var info = null;

		for ( var i = 0; i < self.sheets.length; ++i ) {
			if ( ( typeof ( sheet ) === "number" && self.sheets[ i ].id === sheet ) || ( self.sheets[ i ].name === sheet ) ) {
				info = self.sheets[ i ];
				break;
			}
		}

		if ( info === null ) {
			throw new Error( "Sheet " + sheet + " not found" );
		}

		return {
			filename: info.filename,
			name: info.name,
			id: info.id,
			root: etree.parse( self.archive.file( info.filename ).asText() ).getroot()
		};
	};
	// Load tables for a given sheet
	Workbook.prototype.loadTables = function ( sheet, sheetFilename ) {
		var self = this;

		var sheetDirectory = path.dirname( sheetFilename ),
			sheetName = path.basename( sheetFilename ),
			relsFilename = sheetDirectory + "/" + '_rels' + "/" + sheetName + '.rels',
			relsFile = self.archive.file( relsFilename ),
			tables = []; // [{filename: ..., root: ....}]

		if ( relsFile === null ) {
			return tables;
		}

		var rels = etree.parse( relsFile.asText() ).getroot();

		sheet.findall( "tableParts/tablePart" ).forEach( function ( tablePart ) {
			var relationshipId = tablePart.attrib[ 'r:id' ],
				target = rels.find( "Relationship[@Id='" + relationshipId + "']" ).attrib.Target,
				tableFilename = sheetDirectory + "/" + target,
				tableTree = etree.parse( self.archive.file( tableFilename ).asText() );

			tables.push( {
				filename: tableFilename,
				root: tableTree.getroot()
			} );
		} );

		return tables;
	};

	/**
	 * Generate a new binary .xlsx file
	 */
	Workbook.prototype.generate = function () {
		var self = this;

		// XXX: Getting errors with compression DEFLATE
		return self.archive.generate( {
			base64: false
			/*,
            compression: 'DEFLATE'*/
		} );
	};


	/* Core */
	Workbook.prototype.substitute = function ( sheetNumber, substitutionsData ) {
		var self = this;
		//Contain : name,id,
		var sheet = self.loadSheet( sheetNumber );
		var sheetData = sheet.root.find( "sheetData" ),
			currentRow = null,
			totalRowsInserted = 0,
			namedTables = self.loadTables( sheet.root, sheet.filename ),
			rows = [];

		//Loop over rows
		sheetData.findall( "row" ).forEach( function ( row ) {
			row.attrib.r = currentRow = self.getCurrentRow( row, totalRowsInserted );
			rows.push( row );

			var cells = [],
				cellsInserted = 0,
				newTableRows = [];

			row.findall( "c" ).forEach( function ( cell ) {
				var appendCell = true;
				cell.attrib.r = self.getCurrentCell( cell, currentRow, cellsInserted );


				if ( cell.attrib.t === "s" ) {
					//We look for the key of our string and seek the string value in sharedStrings
					var cellValue = cell.find( 'v' );
					var stringIndex = parseInt( cellValue.text, 10 );
					var string = self.sharedStrings[ stringIndex ];

					if ( string === undefined ) {
						return;
					}

					//Loop Over PlaceHolders
					self.extractPlaceholders( string ).forEach( function ( placeholder ) {
						// Only substitute things for which we have a substitution
						var substitution = substitutionsData[ placeholder.name ],
							newCellsInserted = 0;
						if ( substitution === undefined ) {
							return;
						}
						//We are here, just if we are a match between sharedStrings and passed data
						if ( placeholder.full && placeholder.type === "table" && substitution instanceof Array ) {
							console.log( "We found an array" );

							/*	console.log( placeholder.key );
							var keys = placeholder.key.split( '.' );
							var keySize = keys.length;

							console.log( keySize );*/

							/*if ( keySize == 2 ) {
								newCellsInserted = self.substituteDoubleTable(
									row, newTableRows,
									cells, cell,
									namedTables, substitution, placeholder.key
								);
							}
							else {*/
							newCellsInserted = self.substituteForOsurvey(
								row, newTableRows,
								cells, cell,
								namedTables, substitution, placeholder.key
							);
							//}


							console.log( "newCellsInserted=" + newCellsInserted );

							//Did we insert new columns (array values)?
							if ( newCellsInserted !== 0 ) {
								appendCell = false;
								cellsInserted += newCellsInserted;
								self.pushRight( self.workbook, sheet.root, cell.attrib.r, newCellsInserted );
							}


						}
						else if ( placeholder.full && placeholder.type === "normal" && substitution instanceof Array ) {
							console.log( "We found an normal stuff" );

							appendCell = false; // don't double-insert cells
							newCellsInserted = self.substituteArray(
								cells, cell, substitution
							);

							if ( newCellsInserted !== 0 ) {
								cellsInserted += newCellsInserted;
								self.pushRight( self.workbook, sheet.root, cell.attrib.r, newCellsInserted );
							}
						}
						else {
							console.log( "We found a classic stuff" );
							string = self.substituteScalar( cell, string, placeholder, substitution );
						}
					} );
				}
				//If we insert column, we may not want to keep original cell anymore
				if ( appendCell ) {
					cells.push( cell );
					console.log( "--" );
				}
			} ); //End of cell Loop

			//In case we inserted column we have to re-build the children of the row
			self.replaceChildren( row, cells );

			if ( cellsInserted !== 0 ) {
				self.updateRowSpan( row, cellsInserted );
			}

			//Add inserted rows
			if ( newTableRows.length > 0 ) {
				newTableRows.forEach( function ( row ) {
					rows.push( row );
					++totalRowsInserted;
				} );

				self.pushDown( self.workbook, sheet.root, namedTables, currentRow, newTableRows.length );
			}
		} ); //End of row loop

		//Rebuild children of sheetData
		self.replaceChildren( sheetData, rows );

		//Update placeholders in table column headers
		self.substituteTableColumnHeaders( namedTables, substitutionsData );


		// Write back the modified XML trees
		self.archive.file( sheet.filename, etree.tostring( sheet.root ) );
		self.archive.file( self.workbookPath, etree.tostring( self.workbook ) );
		self.writeSharedStrings();
		self.writeTables( namedTables );

	};



	/* Modified Helpers */


	// Perform a table substitution. May update `newTableRows` and `cells` and change `cell`.
	// Returns total number of new cells inserted on the original row.
	Workbook.prototype.substituteTable = function ( row, newTableRows, cells, cell, namedTables, substitution, key ) {
		var self = this,
			newCellsInserted = 0; // on the original row

		// if no elements, blank the cell, but don't delete it
		if ( substitution.length === 0 ) {
			delete cell.attrib.t;
			self.replaceChildren( cell, [] );
		}
		else {

			var parentTables = namedTables.filter( function ( namedTable ) {
				var range = self.splitRange( namedTable.root.attrib.ref );
				return self.isWithin( cell.attrib.r, range.start, range.end );
			} );

			substitution.forEach( function ( element, idx ) {
				var newRow, newCell,
					newCellsInsertedOnNewRow = 0,
					newCells = [],
					value = element[ key ];


				if ( idx === 0 ) { // insert in the row where the placeholders are

					if ( value instanceof Array ) {
						newCellsInserted = self.substituteArray( cells, cell, value );
					}
					else {
						self.insertCellValue( cell, value );
					}

				}
				else { // insert new rows (or reuse rows just inserted)

					// Do we have an existing row to use? If not, create one.
					if ( ( idx - 1 ) < newTableRows.length ) {
						newRow = newTableRows[ idx - 1 ];
					}
					else {
						newRow = self.cloneElement( row, false );
						newRow.attrib.r = self.getCurrentRow( row, newTableRows.length + 1 );
						newTableRows.push( newRow );
					}

					// Create a new cell
					newCell = self.cloneElement( cell );
					newCell.attrib.r = self.joinRef( {
						row: newRow.attrib.r,
						col: self.splitRef( newCell.attrib.r ).col
					} );

					if ( value instanceof Array ) {
						newCellsInsertedOnNewRow = self.substituteArray( newCells, newCell, value );

						// Add each of the new cells created by substituteArray()
						newCells.forEach( function ( newCell ) {
							newRow.append( newCell );
						} );

						self.updateRowSpan( newRow, newCellsInsertedOnNewRow );
					}
					else {
						self.insertCellValue( newCell, value );

						// Add the cell that previously held the placeholder
						newRow.append( newCell );
					}

					// expand named table range if necessary
					parentTables.forEach( function ( namedTable ) {
						var tableRoot = namedTable.root,
							autoFilter = tableRoot.find( "autoFilter" ),
							range = self.splitRange( tableRoot.attrib.ref );

						if ( !self.isWithin( newCell.attrib.r, range.start, range.end ) ) {
							range.end = self.nextRow( range.end );
							tableRoot.attrib.ref = self.joinRange( range );
							if ( autoFilter !== null ) {
								// XXX: This is a simplification that may stomp on some configurations
								autoFilter.attrib.ref = tableRoot.attrib.ref;
							}
						}
					} );
				}
			} );
		}
		return newCellsInserted;
	};

	// Perform a table substitution. May update `newTableRows` and `cells` and change `cell`.
	// Returns total number of new cells inserted on the original row.
	Workbook.prototype.substituteForOsurvey = function ( row, newTableRows, cells, cell, namedTables, substitution, key ) {
		var self = this,
			newCellsInserted = 0; // on the original row

		// if no elements, blank the cell, but don't delete it
		if ( substitution.length === 0 ) {
			delete cell.attrib.t;
			self.replaceChildren( cell, [] );
		}
		else {

			var parentTables = namedTables.filter( function ( namedTable ) {
				var range = self.splitRange( namedTable.root.attrib.ref );
				return self.isWithin( cell.attrib.r, range.start, range.end );
			} );

			substitution.forEach( function ( element, idx ) {
				var keys = key.split( '.' );
				var keySize = keys.length;
				if ( keySize == 2 && element[ keys[ 0 ] ] instanceof Array ) {
					var elSize = parseInt( element[ keys[ 0 ] ].length - 1 );
					element[ keys[ 0 ] ].forEach( function ( element, idx ) {
						var newRow, newCell,
							newCellsInsertedOnNewRow = 0,
							newCells = [],
							value = element[ keys[ 1 ] ];

						if ( idx === 0 ) {
							if ( value instanceof Array ) {
								newCellsInserted = self.substituteArray( cells, cell, value );
							}
							else {
								self.insertCellValue( cell, value );
							}
						}
						else {
							if ( ( idx - 1 ) < newTableRows.length ) {
								newRow = newTableRows[ idx - 1 ];
							}
							else {
								newRow = self.cloneElement( row, false );
								newRow.attrib.r = self.getCurrentRow( row, newTableRows.length + 1 );
								newTableRows.push( newRow );
							}


							// Create a new cell
							newCell = self.cloneElement( cell );
							newCell.attrib.r = self.joinRef( {
								row: newRow.attrib.r,
								col: self.splitRef( newCell.attrib.r ).col
							} );
							//console.log( newCell._ );


						}


					} );
				}
			} );
		}
		return newCellsInserted;
	};


	/* Helpers */


	// Add a new shared string
	Workbook.prototype.addSharedString = function ( s ) {
		var self = this;

		var idx = self.sharedStrings.length;
		self.sharedStrings.push( s );
		self.sharedStringsLookup[ s ] = idx;

		return idx;
	};

	// Write back the new shared strings list
	Workbook.prototype.writeSharedStrings = function () {
		var self = this;

		var root = etree.parse( self.archive.file( self.sharedStringsPath ).asText() ).getroot(),
			children = root.getchildren();

		root.delSlice( 0, children.length );

		self.sharedStrings.forEach( function ( string ) {
			var si = new etree.Element( "si" ),
				t = new etree.Element( "t" );

			t.text = string;
			si.append( t );
			root.append( si );
		} );

		root.attrib.count = self.sharedStrings.length;
		root.attrib.uniqueCount = self.sharedStrings.length;

		self.archive.file( self.sharedStringsPath, etree.tostring( root ) );
	};
	Workbook.prototype.byString = function ( o, s ) {
		s = s.replace( /\[(\w+)\]/g, '.$1' ); // convert indexes to properties
		s = s.replace( /^\./, '' ); // strip leading dot
		var a = s.split( '.' );
		while ( a.length ) {
			var n = a.shift();
			if ( n in o ) {
				o = o[ n ];
			}
			else {
				return;
			}
		}
		return o;
	};
	// Split a range like "A1:B1" into {start: "A1", end: "B1"}
	Workbook.prototype.splitRange = function ( range ) {
		var split = range.split( ":" );
		return {
			start: split[ 0 ],
			end: split[ 1 ]
		};
	};

	// Clone an element. If `deep` is true, recursively clone children
	Workbook.prototype.cloneElement = function ( element, deep ) {
		var self = this;

		var newElement = etree.Element( element.tag, element.attrib );
		newElement.text = element.text;
		newElement.tail = element.tail;

		if ( deep !== false ) {
			element.getchildren().forEach( function ( child ) {
				newElement.append( self.cloneElement( child, deep ) );
			} );
		}

		return newElement;
	};
	// Replace a shared string with a new one at the same index. Return the
	// index.
	Workbook.prototype.replaceString = function ( oldString, newString ) {
		var self = this;

		var idx = self.sharedStringsLookup[ oldString ];
		if ( idx === undefined ) {
			idx = self.addSharedString( newString );
		}
		else {
			self.sharedStrings[ idx ] = newString;
			delete self.sharedStringsLookup[ oldString ];
			self.sharedStringsLookup[ newString ] = idx;
		}

		return idx;
	};

	// Get the number of a shared string, adding a new one if necessary.
	Workbook.prototype.stringIndex = function ( s ) {
		var self = this;

		var idx = self.sharedStringsLookup[ s ];
		if ( idx === undefined ) {
			idx = self.addSharedString( s );
		}
		return idx;
	};
	// Turn a value of any type into a string
	Workbook.prototype.stringify = function ( value ) {
		var self = this;

		if ( value instanceof Date ) {
			return value.toISOString();
		}
		else if ( typeof ( value ) === "number" || typeof ( value ) === "boolean" ) {
			return Number( value ).toString();
		}
		else {
			return String( value ).toString();
		}
	};

	// Insert a substitution value into a cell (c tag)
	Workbook.prototype.insertCellValue = function ( cell, substitution ) {
		var self = this;

		var cellValue = cell.find( "v" ),
			stringified = self.stringify( substitution );

		if ( typeof ( substitution ) === "number" ) {
			cell.attrib.t = "n";
			cellValue.text = stringified;
		}
		else if ( typeof ( substitution ) === "boolean" ) {
			cell.attrib.t = "b";
			cellValue.text = stringified;
		}
		else if ( substitution instanceof Date ) {
			cell.attrib.t = "d";
			cellValue.text = stringified;
		}
		else {
			cell.attrib.t = "s";
			cellValue.text = Number( self.stringIndex( stringified ) ).toString();
		}

		return stringified;
	};
	// Perform substitution of a single value
	Workbook.prototype.substituteScalar = function ( cell, string, placeholder, substitution ) {
		var self = this;

		if ( placeholder.full && typeof ( substitution ) === "string" ) {
			self.replaceString( string, substitution );
		}

		if ( placeholder.full ) {
			return self.insertCellValue( cell, substitution );
		}
		else {
			var newString = string.replace( placeholder.placeholder, self.stringify( substitution ) );
			cell.attrib.t = "s";
			self.replaceString( string, newString );
			return newString;
		}

	};

	// Write back possibly-modified tables
	Workbook.prototype.writeTables = function ( tables ) {
		var self = this;

		tables.forEach( function ( namedTable ) {
			self.archive.file( namedTable.filename, etree.tostring( namedTable.root ) );
		} );
	};
	// Calculate the current row based on a source row and a number of new rows
	// that have been inserted above
	Workbook.prototype.getCurrentRow = function ( row, rowsInserted ) {
		return parseInt( row.attrib.r, 10 ) + rowsInserted;
	};
	// Calculate the current cell based on asource cell, the current row index,
	// and a number of new cells that have been inserted so far
	Workbook.prototype.getCurrentCell = function ( cell, currentRow, cellsInserted ) {
		var self = this;

		var colRef = self.splitRef( cell.attrib.r ).col,
			colNum = self.charToNum( colRef );

		return self.joinRef( {
			row: currentRow,
			col: self.numToChar( colNum + cellsInserted )
		} );
	};
	// Split a reference into an object with keys `row` and `col` and,
	// optionally, `table`, `rowAbsolute` and `colAbsolute`.
	Workbook.prototype.splitRef = function ( ref ) {
		var match = ref.match( /(?:(.+)!)?(\$)?([A-Z]+)(\$)?([0-9]+)/ );
		return {
			table: match[ 1 ] || null,
			colAbsolute: Boolean( match[ 2 ] ),
			col: match[ 3 ],
			rowAbsolute: Boolean( match[ 4 ] ),
			row: parseInt( match[ 5 ], 10 )
		};
	};
	// Turn a reference like "AA" into a number like 27
	Workbook.prototype.charToNum = function ( str ) {
		var num = 0;
		for ( var idx = str.length - 1, iteration = 0; idx >= 0; --idx, ++iteration ) {
			var thisChar = str.charCodeAt( idx ) - 64, // A -> 1; B -> 2; ... Z->26
				multiplier = Math.pow( 26, iteration );
			num += multiplier * thisChar;
		}
		return num;
	};
	// Turn a number like 27 into a reference like "AA"
	Workbook.prototype.numToChar = function ( num ) {
		var str = "";


		for ( var i = 0; num > 0; ++i ) {
			var remainder = num % 26,
				charCode = remainder + 64;
			num = ( num - remainder ) / 26;

			// Compensate for the fact that we don't represent zero, e.g. A = 1, Z = 26, but AA = 27
			if ( remainder === 0 ) { // 26 -> Z
				charCode = 90;
				--num;
			}

			str = String.fromCharCode( charCode ) + str;
		}

		return str;
	};
	// Join an object with keys `row` and `col` into a single reference string
	Workbook.prototype.joinRef = function ( ref ) {
		return ( ref.table ? ref.table + "!" : "" ) +
			( ref.colAbsolute ? "$" : "" ) +
			ref.col.toUpperCase() +
			( ref.rowAbsolute ? "$" : "" ) +
			Number( ref.row ).toString();
	};

	// Return a list of tokens that may exist in the string.
	// Keys are: `placeholder` (the full placeholder, including the `${}`
	// delineators), `name` (the name part of the token), `key` (the object key
	// for `table` tokens), `full` (boolean indicating whether this placeholder
	// is the entirety of the string) and `type` (one of `table` or `cell`)
	Workbook.prototype.extractPlaceholders = function ( string ) {
		// Yes, that's right. It's a bunch of brackets and question marks and stuff.
		var re = /\${(?:(.+?):)?(.+?)(?:\.(.+?))?}/g;

		var match = null,
			matches = [];
		while ( ( match = re.exec( string ) ) !== null ) {
			matches.push( {
				placeholder: match[ 0 ],
				type: match[ 1 ] || 'normal',
				name: match[ 2 ],
				key: match[ 3 ],
				full: match[ 0 ].length === string.length
			} );
		}

		return matches;
	};
	// Replace all children of `parent` with the nodes in the list `children`
	Workbook.prototype.replaceChildren = function ( parent, children ) {
		parent.delSlice( 0, parent.len() );
		children.forEach( function ( child ) {
			parent.append( child );
		} );
	};

	// Adjust the row `spans` attribute by `cellsInserted`
	Workbook.prototype.updateRowSpan = function ( row, cellsInserted ) {
		if ( cellsInserted !== 0 && row.attrib.spans ) {
			var rowSpan = row.attrib.spans.split( ':' ).map( function ( f ) {
				return parseInt( f, 10 );
			} );
			rowSpan[ 1 ] += cellsInserted;
			row.attrib.spans = rowSpan.join( ":" );
		}
	};
	// Look for any merged cell, named table or named range definitions below
	// `currentRow` and push down by `numRows` (used when rows are inserted).
	Workbook.prototype.pushDown = function ( workbook, sheet, tables, currentRow, numRows ) {
		var self = this;

		// Update merged cells below this row
		sheet.findall( "mergeCells/mergeCell" ).forEach( function ( mergeCell ) {
			var mergeRange = self.splitRange( mergeCell.attrib.ref ),
				mergeStart = self.splitRef( mergeRange.start ),
				mergeEnd = self.splitRef( mergeRange.end );

			if ( mergeStart.row > currentRow ) {
				mergeStart.row += numRows;
				mergeEnd.row += numRows;

				mergeCell.attrib.ref = self.joinRange( {
					start: self.joinRef( mergeStart ),
					end: self.joinRef( mergeEnd ),
				} );

			}
		} );

		// Update named tables below this row
		tables.forEach( function ( table ) {
			var tableRoot = table.root,
				tableRange = self.splitRange( tableRoot.attrib.ref ),
				tableStart = self.splitRef( tableRange.start ),
				tableEnd = self.splitRef( tableRange.end );

			if ( tableStart.row > currentRow ) {
				tableStart.row += numRows;
				tableEnd.row += numRows;

				tableRoot.attrib.ref = self.joinRange( {
					start: self.joinRef( tableStart ),
					end: self.joinRef( tableEnd ),
				} );

				var autoFilter = tableRoot.find( "autoFilter" );
				if ( autoFilter !== null ) {
					// XXX: This is a simplification that may stomp on some configurations
					autoFilter.attrib.ref = tableRoot.attrib.ref;
				}
			}

		} );

		// Named cells/ranges
		workbook.findall( "definedNames/definedName" ).forEach( function ( name ) {
			var ref = name.text;

			if ( self.isRange( ref ) ) {
				var namedRange = self.splitRange( ref ),
					namedStart = self.splitRef( namedRange.start ),
					namedEnd = self.splitRef( namedRange.end );

				if ( namedStart.row > currentRow ) {
					namedStart.row += numRows;
					namedEnd.row += numRows;

					name.text = self.joinRange( {
						start: self.joinRef( namedStart ),
						end: self.joinRef( namedEnd ),
					} );

				}
			}
			else {
				var namedRef = self.splitRef( ref ),
					namedCol = self.charToNum( namedRef.col );

				if ( namedRef.row > currentRow ) {
					namedRef.row += numRows;
					name.text = self.joinRef( namedRef );
				}
			}

		} );
	};

	// Perform substitution in table headers
	Workbook.prototype.substituteTableColumnHeaders = function ( tables, substitutions ) {
		var self = this;

		tables.forEach( function ( table ) {
			var root = table.root,
				columns = root.find( "tableColumns" ),
				autoFilter = root.find( "autoFilter" ),
				tableRange = self.splitRange( root.attrib.ref ),
				idx = 0,
				inserted = 0,
				newColumns = [];

			columns.findall( "tableColumn" ).forEach( function ( col ) {
				++idx;
				col.attrib.id = Number( idx ).toString();
				newColumns.push( col );

				var name = col.attrib.name;

				self.extractPlaceholders( name ).forEach( function ( placeholder ) {
					var substitution = substitutions[ placeholder.name ];
					if ( substitution === undefined ) {
						return;
					}

					// Array -> new columns
					if ( placeholder.full && placeholder.type === "normal" && substitution instanceof Array ) {
						substitution.forEach( function ( element, i ) {
							var newCol = col;
							if ( i > 0 ) {
								newCol = self.cloneElement( newCol );
								newCol.attrib.id = Number( ++idx ).toString();
								newColumns.push( newCol );
								++inserted;
								tableRange.end = self.nextCol( tableRange.end );
							}
							newCol.attrib.name = self.stringify( element );
						} );
						// Normal placeholder
					}
					else {
						name = name.replace( placeholder.placeholder, self.stringify( substitution ) );
						col.attrib.name = name;
					}
				} );
			} );

			self.replaceChildren( columns, newColumns );

			// Update range if we inserted columns
			if ( inserted > 0 ) {
				columns.attrib.count = Number( idx ).toString();
				root.attrib.ref = self.joinRange( tableRange );
				if ( autoFilter !== null ) {
					// XXX: This is a simplification that may stomp on some configurations
					autoFilter.attrib.ref = self.joinRange( tableRange );
				}
			}

		} );
	};
	return Workbook;
} )();