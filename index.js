var XlsxTemplate = require('xlsx-template');
var fs = require('fs');
var path = require('path');

// Load an XLSX file into memory
fs.readFile(path.join(__dirname, 'templates', 'Teste.xlsx'), function(err, data) {

    // Create a template
    var template = new XlsxTemplate(data);

    // Replacements take place on first sheet
    var sheetNumber = 1;

    // Set up some placeholder values matching the placeholders in the template
    var values = {"contas": [{"receitas":1, "despesas":22}, {"receitas":55, "despesas":65}]};

    // Perform substitution
    template.substitute(sheetNumber, values);

    // Get binary data
    var data = template.generate({type: 'uint8array'});

    fs.writeFileSync(path.join(__dirname, 'templates', 'out.xlsx'), data)

});