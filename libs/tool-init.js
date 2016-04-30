var X = XLSX;

$('#refresh').click(function(){
    process_wb(wb);
});

var drop = document.getElementById('drop');

function fixdata(data) {
    var o = "", l = 0, w = 10240;
    for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
    o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
    return o;
}

function process_wb(wb) {
    $('#message').text('處理中');
    var output = "";
    output = transfer_workbook(wb);
    $('#result').val(output);
}

function transfer_workbook(workbook) {
    var warnings = [];
    var char_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    var char_to_int = function(char){
	var r = 0;
	for (var i = 0; i < char.length; i ++) {
	    r *= 26;
	    r += char_list.indexOf(char[i]) + 1;
	}
	return r;
    };

    var int_to_char = function(i){
	var c = '';
	i --;
	do {
	    if (c == '') {
		c = char_list[i % 26];
	    } else {
		c = char_list[(i - 1) % 26] + c;
	    }
	    i = Math.floor(i / 26);
	} while (i > 0);
	return c;
    }

    ret = {};
    for (var tab in workbook.Sheets) {
        ret[tab] = {};
	if ('undefined' !== typeof(workbook.Sheets[tab]['!ref'])) {
	    var ref = workbook.Sheets[tab]['!ref'];
	    var match = ref.match(/([A-Z]*)([0-9]*):([A-Z]*)([0-9]*)/);
	    range = {
e: {
r: parseInt(match[4]),
   c: char_to_int(match[3])
   },
s: {
r: parseInt(match[2]),
   c: char_to_int(match[1])
   }
	    };
	} else {
	    range = workbook.Sheets[tab]['!range'];
	}
        ret[tab].width = 0;
        ret[tab].data = [];
	for (var j = range.s.r; j <= range.e.r; j ++) {
	    var row = [];
	    isempty = -1;
	    for (var k = range.s.c; k <= range.e.c; k ++){
		if ('undefined' === typeof(workbook.Sheets[tab][int_to_char(k) + j])) {
		    row.push(null);
		} else {
                    if (workbook.Sheets[tab][int_to_char(k) + j].w == '#REF!') {
                        warnings.push("分頁" + tab + " 的 " + int_to_char(k) + j + "格有錯誤的公式");
                    }
		    row.push(workbook.Sheets[tab][int_to_char(k) + j].w);
		    isempty = row.length;
		}
	    }
	    if (!excel_parse_options["ignore-empty-line"] || isempty > 0) {
                if (excel_parse_options["ignore-line-tail-null"]) {
                    row = row.slice(0, isempty);
                }
		ret[tab].data.push(row);
                ret[tab].width = Math.max(ret[tab].width, row.length);
	    }
	}
        if (excel_parse_options["with-merge-cells"] && 'undefined' !== typeof(workbook.Sheets[tab]['!merges'])) {
            for ( var i = 0; i < workbook.Sheets[tab]['!merges'].length; i ++ ) {
                var merge = workbook.Sheets[tab]['!merges'][i];
                for (var c = merge.s.c; c <= merge.e.c; c++) {
                    for (var r = merge.s.r; r <= merge.e.r; r ++) {
                        if (c == merge.s.c && r == merge.s.r) {
                            continue;
                        }
                        ret[tab].data[r][c] = {'type' : 'merge', 'from' : [merge.s.r , merge.s.c]};
                    }
                }
            }
        }
        ret[tab].height = ret[tab].data.length;
    }

    ret = main(ret, warnings);
    csv_array = ret[0];
    warnings = ret[1];

    split = $('input[name="split"]:checked').val();
    result = '';
    for (var line_no = 0; line_no < csv_array.length; line_no ++) {
        if (split == ',') {
            result += csv_array[line_no].map(function(v) {
                if ('undefined' === typeof(v) || v === null) {
                    return v;
                }
                if (v.indexOf('"')) {
                    v = v.replace(/"/g, '""');
                }
                if (v.indexOf('"') >= 0 || v.indexOf("\n") >= 0 || v.indexOf(",") >= 0) {
                    v = '"' + v + '"';
                }
                return v;
            }).join(',') + "\n";
        } else if (split == 'tab') {
            result += csv_array[line_no].map(function(v){
                if ('undefined' === typeof(v) || v === null) {
                    return v;
                }
                if (v.indexOf("\n") >= 0) {
                    v = v.replace(/\n/mg, ' ');
                }
                return v;
            }).join("\t") + "\n";
        }
    }
    $('#message').text("處理完成\n" + warnings.join("\n"));
    return result;
}

var wb;
function handleDrop(e) {
    $('#message').text('處理中');
    e.stopPropagation();
    e.preventDefault();
    var files = e.dataTransfer.files;
    var f = files[0];
    {
        var reader = new FileReader();
        var name = f.name;
        reader.onload = function(e) {
            var data = e.target.result;
            var arr = fixdata(data);
            wb = X.read(btoa(arr), {type: 'base64'});
            process_wb(wb);
        };
        reader.readAsArrayBuffer(f);
    }
}

function handleDragover(e) {
    e.stopPropagation();
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
}

if(drop.addEventListener) {
    drop.addEventListener('dragenter', handleDragover, false);
    drop.addEventListener('dragover', handleDragover, false);
    drop.addEventListener('drop', handleDrop, false);
}

var xlf = document.getElementById('xlf');
function handleFile(e) {
    $('#message').text('處理中');
	var files = e.target.files;
	var f = files[0];
	{
		var reader = new FileReader();
		var name = f.name;
		reader.onload = function(e) {
			var data = e.target.result;
			var wb;
			var arr = fixdata(data);
			wb = X.read(btoa(arr), {type: 'base64'});
			process_wb(wb);
		};
		reader.readAsArrayBuffer(f);
	}
}

if(xlf.addEventListener) xlf.addEventListener('change', handleFile, false);

$('#download-form').submit(function(e){
	e.preventDefault();
        var a = document.createElement('a');
        var blob = new Blob([$('#result').val()], {'type': 'octet/stream'});
        a.href = window.URL.createObjectURL(blob);
        a.download = $('#download-file').val();
        a.click();
});
