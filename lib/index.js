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

	/**
	 * Create a new workbook. Either pass the raw data of a .xlsx file,
	 * or call `loadTemplate()` later.
	 */
	var Workbook = function ( data ) {
		var self = this;

		self.archive = null;
		self.sharedStrings = [];
		self.sharedStringsLookup = {};

		if ( data ) {
			self.loadTemplate( data );
		}
	};

	/**
	 * Load a .xlsx file from a byte array.
	 */
	Workbook.prototype.loadTemplate = function ( data ) {

	};

	/**
	 * Interpolate values for the sheet with the given number (1-based) or
	 * name (if a string) using the given substitutions (an object).
	 */
	Workbook.prototype.substitute = function ( sheetName, substitutions ) {

	};

	/**
	 * Generate a new binary .xlsx file
	 */
	Workbook.prototype.generate = function () {

	};

	// Helpers

	// Write back the new shared strings list
	Workbook.prototype.writeSharedStrings = function () {

	};

	// Add a new shared string
	Workbook.prototype.addSharedString = function ( s ) {

	};

	// Get the number of a shared string, adding a new one if necessary.
	Workbook.prototype.stringIndex = function ( s ) {

	};

	// Replace a shared string with a new one at the same index. Return the
	// index.
	Workbook.prototype.replaceString = function ( oldString, newString ) {

	};

	// Get a list of sheet ids, names and filenames
	Workbook.prototype.loadSheets = function ( prefix, workbook, workbookRels ) {

	};

	// Get sheet a sheet, including filename and name
	Workbook.prototype.loadSheet = function ( sheet ) {

	};

	// Load tables for a given sheet
	Workbook.prototype.loadTables = function ( sheet, sheetFilename ) {

	};

	// Write back possibly-modified tables
	Workbook.prototype.writeTables = function ( tables ) {

	};

	// Perform substitution in table headers
	Workbook.prototype.substituteTableColumnHeaders = function ( tables, substitutions ) {

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


	return Workbook;
} )();