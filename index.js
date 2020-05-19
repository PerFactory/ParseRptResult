const fs = require('fs');
const yargs = require('yargs');
const xl = require('excel4node');

const yargv = yargs
    .option('input', {
        alias: 'i',
        description: 'The html report folder path',
        type: 'string',
    })
	.option('output', {
        alias: 'o',
        description: 'The path for resutl excel file path',
        type: 'text',
    })
    .help()
    .alias('help', 'h')
    .argv;


if(!yargv.input){
	console.log('Input folder path required!');
	return;
}

// read & init configuration from JSON file
const config = JSON.parse(fs.readFileSync('config.json', 'utf8'));

console.log('Reading data from ' + yargv.input);

// read requests_map file save into a json letiable request_map
const req_map_path = yargv.input + '/datas/requests_map.js';
const requests_map_raw = fs.readFileSync(req_map_path, 'utf8');
const requests_map = JSON.parse(requests_map_raw.substring(49, requests_map_raw.length - 3));

// find the request js file from request_map (both header and data)
let tran_file;
let page_file;
let test_detail_file;

Object.keys(requests_map).forEach( (k) => {
		if(k.indexOf(config.transaction_key) > 0){
			tran_file = yargv.input + '/datas/' + requests_map[k] + '.js';
		}
		if(k.indexOf(config.page_key) > 0){
			page_file = yargv.input + '/datas/' + requests_map[k] + '.js';
		}
		if(k.indexOf(config.test_detail_key) > 0){
			test_detail_file = yargv.input + '/datas/' + requests_map[k] + '.js';
		}
	});


// create json list with header & data

let transactions = [{"col_0":"Transaction Name", "Col_1":"Hits", "col_2":"Average", "col_3":"Min", "col_4":"Max", "Col_5":"75 Percentile", "col_6":"90 Percentile"}];
let tran_arr = [];

const tran_data_raw = fs.readFileSync(tran_file, 'utf8');
const tran_data = JSON.parse(tran_data_raw.substring(tran_data_raw.indexOf('return ') + 7, tran_data_raw.length - 3));

for ( let obj of tran_data.groups[0].instances ) {
	tran_arr.push({"col_0":obj.name, "Col_1": {"value":checkValue(obj.counters[0]), "highlight": false}, 
		"col_2":hightlightTran(config.transaction_sla, obj.name, obj.counters[1]), 
		"col_3":hightlightTran(config.transaction_sla, obj.name, obj.counters[2]), 
		"col_4":hightlightTran(config.transaction_sla, obj.name, obj.counters[3]), 
		"Col_5":hightlightTran(config.transaction_sla, obj.name, obj.counters[4]), 
		"col_6":hightlightTran(config.transaction_sla, obj.name, obj.counters[5])
	});
}

tran_arr.sort((a,b)=> (a.col_0 > b.col_0) ? 1 : -1);

Array.prototype.push.apply(transactions,tran_arr);

let pages = [{"col_1":"Request Name", "Col_2":"Hits", "col_3":"Average", "col_4":"Min", "col_5":"Max", "Col_6":"75 Percentile", "col_7":"90 Percentile"}];

const pages_data_raw = fs.readFileSync(page_file, 'utf8');
const pages_data = JSON.parse(pages_data_raw.substring(pages_data_raw.indexOf('return ') + 7, pages_data_raw.length - 3));

for ( let obj of pages_data.groups[0].instances ) {
	let page_name = obj.name;
	for( let page of obj.groups[0].instances){
		let page_obj = {"Col_1":page.name, "Col_2":checkValue(page.counters[0]), 
			"col_2":convertToSecond(page.counters[2]), 
			"col_3":convertToSecond(page.counters[3]), 
			"col_4":convertToSecond(page.counters[4]), 
			"Col_5":convertToSecond(page.counters[5]), 
			"col_6":convertToSecond(page.counters[6])};
		pages.push(page_obj);
	}
}

// write json list into excel
let wb = CreateWorkbook();

const header_style= wb.createStyle({
	alignment: {
		horizontal: 'center',
		vertical: 'center'
	},
	font: {
		bold: true,
		color: '#FFFFFF',
		size: 14
	},
	border: {
		left: {
			style:'thin',
			color: '#000000'
		},
		right: {
			style:'thin',
			color: '#000000'
		},
		top: {
			style:'thin',
			color: '#000000'
		},
		bottom: {
			style:'thin',
			color: '#000000'
		},
	},
	fill: {
		type: 'pattern',
		patternType: 'solid',
		fgColor: '#203764'
	}
});

const cell_style= wb.createStyle({
	border: {
		left: {
			style:'thin',
			color: '#000000'
		},
		right: {
			style:'thin',
			color: '#000000'
		},
		top: {
			style:'thin',
			color: '#000000'
		},
		bottom: {
			style:'thin',
			color: '#000000'
		},
	}
});

const highlight_cell_style= wb.createStyle({
	font: {
		bold: false,
		color: '#9c0006',
		size: 11
	},
	border: {
		left: {
			style:'thin',
			color: '#000000'
		},
		right: {
			style:'thin',
			color: '#000000'
		},
		top: {
			style:'thin',
			color: '#000000'
		},
		bottom: {
			style:'thin',
			color: '#000000'
		},
	},
	fill: {
		type: 'pattern',
		patternType: 'solid',
		fgColor: '#ffc7ce'
	}
});

// add test details sheet
let details = getTestDetails(test_detail_file);
const testName = details.testName;
const startTime = details.startTime;
const duration = details.startTime + details.timeRanges.run.duration;

let detail_ws = wb.addWorksheet('Test Details', ws_options);

detail_ws.cell(2,2).string("Test Name").style(header_style);
detail_ws.cell(2,3).string(testName).style(cell_style);

detail_ws.cell(3,2).string("Start Date time").style(header_style);
detail_ws.cell(3,3).string(new Date(details.startTime).toLocaleString()).style(cell_style);

detail_ws.cell(4,2).string("End Date time").style(header_style);
detail_ws.cell(4,3).string(new Date(details.startTime + details.timeRanges.run.duration).toLocaleString()).style(cell_style);

detail_ws.column(2).setWidth(20);
detail_ws.column(3).setWidth(30);

// add transaction sheet & data
let tran_ws = wb.addWorksheet('Transaction Details', ws_options);
let row = 2, col = 1;

for(let tran of transactions){
	Object.keys(tran).forEach((k) => {
		if(row == 2) {// header
			tran_ws.cell(row, col).string(tran[k]).style(header_style);
		}else{
			if (col == 1)
				tran_ws.cell(row, col).string(tran[k]).style(cell_style);
			else{
				if(tran[k].highlight)
					tran_ws.cell(row, col).number(tran[k].value).style(highlight_cell_style);
				else
					tran_ws.cell(row, col).number(tran[k].value).style(cell_style);
			}
		}
		col ++;
	});
	col = 1;
	row ++;
}

tran_ws.row(2).filter({ firstRow: 2, firstColumn: 1, lastRow: row, lastColumn: 7});
tran_ws.column(1).setWidth(75);
tran_ws.column(4).hide();
tran_ws.column(5).hide();
tran_ws.column(6).hide();

// add page sheet & data
let page_ws = wb.addWorksheet('Page Details', ws_options);
row = 2;
col = 1;

for(let page of pages){
	Object.keys(page).forEach((k) => {
		if(row == 2) {// header
			page_ws.cell(row, col).string(page[k]).style(header_style);
		}else{
			if (col == 1)
				page_ws.cell(row, col).string(page[k]).style(cell_style);
			else
				page_ws.cell(row, col).number(page[k]).style(cell_style);
		}
		col ++;
	});
	col = 1;
	row ++;
}

page_ws.row(2).filter({ firstRow: 2, firstColumn: 1, lastRow: row-1, lastColumn: 7});
page_ws.column(1).setWidth(75);
page_ws.column(4).hide();
page_ws.column(5).hide();
page_ws.column(6).hide();

// determine output path and flush the excel file
outputFile = yargv.output + '/' + 'result_' + new Date(details.startTime).toISOString().replace(/:/g, '_') + '_temp.xlsx';
wb.write(outputFile, (err, stat) => { if (err) { console.error(err); } else{ console.log('Saved to ' + outputFile); } });

function CreateWorkbook(){
	return new xl.Workbook({
			jszip: {
				compression: 'DEFLATE',
			},
			defaultFont: {
				size: 11,
				name: 'Calibri',
				color: '000000',
			},
			dateFormat: 'd/m/yy hh:mm:ss',
			workbookView: {
				activeTab: 1, // Specifies an unsignedInt that contains the index to the active sheet in this book view.
				autoFilterDateGrouping: true, // Specifies a boolean value that indicates whether to group dates when presenting the user with filtering options in the user interface.
				//firstSheet: 1, // Specifies the index to the first sheet in this book view.
				minimized: false, // Specifies a boolean value that indicates whether the workbook window is minimized.
				showHorizontalScroll: true, // Specifies a boolean value that indicates whether to display the horizontal scroll bar in the user interface.
				showSheetTabs: true, // Specifies a boolean value that indicates whether to display the sheet tabs in the user interface.
				showVerticalScroll: true, // Specifies a boolean value that indicates whether to display the vertical scroll bar.
				tabRatio: 600, // Specifies ratio between the workbook tabs bar and the horizontal scroll bar.
				visibility: 'visible', // Specifies visible state of the workbook window. ('hidden', 'veryHidden', 'visible') (§18.18.89)
				windowHeight: 17620, // Specifies the height of the workbook window. The unit of measurement for this value is twips.
				windowWidth: 28800, // Specifies the width of the workbook window. The unit of measurement for this value is twips..
				xWindow: 0, // Specifies the X coordinate for the upper left corner of the workbook window. The unit of measurement for this value is twips.
				yWindow: 440, // Specifies the Y coordinate for the upper left corner of the workbook window. The unit of measurement for this value is twips.
			},
			logLevel: 5, // 0 - 5. 0 suppresses all logs, 1 shows errors only, 5 is for debugging
			author: 'Shiba Prasad Swain', // Name for use in features such as comments
		});
}

var ws_options = {
  margins: {
    left: 1.5,
    right: 1.5,
  },
};

function convertToSecond(val){
	if (Number.isNaN(Number.parseFloat(val))) {
		return 0;
	}
	return parseFloat((val/1000).toFixed(3));
}

function checkValue(val){
	if (Number.isNaN(Number.parseFloat(val))) {
		return 0;
	}
	return val;
}

function getTestDetails(path){
	const raw_data = fs.readFileSync(path, 'utf8');
	return JSON.parse(raw_data.substring(raw_data.indexOf('return ') + 7, raw_data.length - 3));
}

function hightlightTran(tranConf, tranName, value){
	secValue = convertToSecond(value);
	
	let sla = tranConf["global"];
	// find the value for transaction from config
	if(tranConf[tranName])
		sla = tranConf[tranName];
	
	let highlight = false;
	// check if value is greater than config value
	if(secValue > sla)
		highlight = true;
	
	// then return the object
	return {"value": secValue, "highlight": highlight};
}
