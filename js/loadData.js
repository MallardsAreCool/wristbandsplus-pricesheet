var excelRows, XLSheets, headerCell, cell, test; //Exposed varibles for testing
test = 1;

//Load File
var url = 'data-sheet/Price List.xlsx?_=' + new Date().getTime();
var oReq = new XMLHttpRequest();
oReq.open('GET', url, true);
oReq.responseType = 'arraybuffer';

oReq.onload = function (e) {
	var arraybuffer = oReq.response;
	var data = new Uint8Array(arraybuffer);
	var arr = new Array();
	for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
	var bstr = arr.join('');

	var workbook = XLSX.read(bstr, {
		type: 'binary',
	});

	//Load Nav
	XLSheets = workbook.SheetNames;

	let types = ['Lanyards', 'Fabric', 'Tyvek', 'Vinyl', 'Silicone', 'Clothing', 'Promotional'];

	for (let i = 0; i < types.length; i++) {
		let tempDiv = document.createElement('div');
		tempDiv.className = 'XLMainBtn';
		let innerDiv = document.createElement('div');
		innerDiv.innerHTML = types[i];
		let dropdown = document.createElement('div');
		dropdown.id = 'dropdown';

		document.querySelectorAll('.XLNav')[0].appendChild(tempDiv);
		tempDiv.appendChild(innerDiv);
		tempDiv.appendChild(dropdown);
	}

	let buttons = document.querySelectorAll('#dropdown');
	for (let i = 0; i < XLSheets.length; i++) {
		var headerButton = document.createElement('button');
		headerButton.className = 'XLNavBtn';
		headerButton.id = i;

		for (let j = 0; j < types.length; j++) {
			if (XLSheets[i].includes(types[j])) {
				headerButton.innerHTML = XLSheets[i].replace(types[j], '');
				buttons[j].appendChild(headerButton);
				break;
			}
			headerButton.innerHTML = XLSheets[i];
			buttons[buttons.length - 1].appendChild(headerButton);
		}
	}

	let navButtons = document.querySelectorAll('.XLNavBtn');
	for (let i = 0; i < navButtons.length; i++) {
		navButtons[i].addEventListener('click', function () {
			let title = this.parentElement.parentElement.innerHTML.split(' ')[0].split('div')[1].substr(1).slice(0, -2);
			loadSheet(this.id, title + ' ' + this.innerHTML);
		});
	}

	function loadSheet(sheetPage, pageName) {
		excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[XLSheets[sheetPage]]);

		var tableHead = document.createElement('div');
		tableHead.className = 'XLTable';
		let counter = 0;

		//WHY 2???? 1 returns less. 0 returns IDs or Length? excelRows is the row number -> 0 = 1, 1 = 2, ect...
		for (var i in excelRows[0]) {
			let XLColumn = document.createElement('div');
			XLColumn.className = 'XLColumn';
			headerCell = document.createElement('div');
			if (i.includes('__EMPTY')) {
				headerCell.innerHTML = 'Blank';
				headerCell.className = 'XLHeader blank';
			} else {
				headerCell.innerHTML = i;
				XLColumn.className += ` ${i}`;
				headerCell.className = 'XLHeader';
			}
			if (i != 'ID') {
				XLColumn.appendChild(headerCell);
				tableHead.appendChild(XLColumn);
			}

			for (var k = 0; k < excelRows.length; k++) {
				cell = document.createElement('div');
				cell.innerHTML = excelRows[k][i];
				cell.className = 'XLInner';
				if (cell.innerHTML.includes('undefined') || cell.innerHTML == ' ' || cell.innerHTML == '') {
					cell.innerHTML = 'Blank';
					cell.className += ' blank';
				}
				cell.className += ` ${k}`;
				XLColumn.appendChild(cell);
			}
		}

		let titleDiv = document.createElement('div');
		titleDiv.className = 'XLTitle';
		titleDiv.innerHTML = pageName;

		var dvExcel = document.getElementById('dvExcel');
		dvExcel.innerHTML = '';
		dvExcel.appendChild(titleDiv);
		dvExcel.appendChild(tableHead);

		prepColumns();
	}
};

oReq.send();
