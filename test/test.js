var xlsexport = require( '../lib' ),
	fs = require( 'fs' ),
	path = require( 'path' ),
	zip = require( 'node-zip' ),
	etree = require( 'elementtree' );
// Load an XLSX file into memory
/*
fs.readFile( path.join( __dirname, 'tpl', 'ExcelB1.xlsx' ), function ( err, data ) {

	// Create a template
	var template = new xlsexport( data );
	// Replacements take place on first sheet
	var sheetNumber = 1;
	// Set up some placeholder values matching the placeholders in the template
	var values = {
		test: "test",
		title: "Uber Test",
		extractDate: new Date(),
		dates: new Date( "2013-06-01" ),
		people: [ {
			name: "John Smith",
			age: 20
		}, {
			name: "Bob Johnson",
			age: 22
		} ]
	};
	// Perform substitution
	template.substitute( sheetNumber, values );

	// Get binary data
	var data_substituted = template.generate();

	try {
		fs.writeFileSync( 'tpl/output/ExcelB1.xlsx', data_substituted, 'binary' );
	}
	catch ( e ) {
		console.error( "Erreur:" + e.message );
	}

	finally {
		/*res.setHeader( 'content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' );
		res.writeHead( 200 );
		res.write( data_substituted, "binary" );

		//res.write( "Génération du rapport réussie. Le fichier viens de vous être envoyé." );
		res.end();*/
//}
// ...

//} ); * /

fs.readFile( path.join( __dirname, 'tpl', 't1.xlsx' ), function ( err, data ) {

	var template = new xlsexport( data );
	var sheetNumber = 1;
	var values = {
		questions: [ {
			formulation: "Quelle est la couleur ?",
			type: "Choix unique",
			answers: [ {
				nbRep: 1,
				formulation: "Bleu"
			}, {
				nbRep: 2,
				formulation: "Rouge"
			} ]
		}, {
			formulation: "Qui est il ?",
			type: "Choix unique",
			answers: [ {
				nbRep: 1,
				formulation: "Qqun"
			}, {
				nbRep: 2,
				formulation: "Personne"
			} ]
		} ]

	};
	template.substitute( sheetNumber, values );
	var data_substituted = template.generate();

	try {
		fs.writeFileSync( 'tpl/output/t1_out.xlsx', data_substituted, 'binary' );
	}
	catch ( e ) {
		console.error( "Erreur:" + e.message );
	}

} );