var wb = XLSX.utils.book_new();

wb.SheetNames.push("Download Vendor Sheet");

var ws_data = [
        ['Name'],
        [],
        ['GID'],
        [],
        ['UID'],
        [],
        ['Email Alias']
];
var ws_content_data = [{
        custodiantype: "external",
        displayname: "John P Patterson",
        globalid: "10001",
        mail: "cn=john p patterson/ou=am/o=lly;cn=john p patterson/ou=am/o=lly@lilly;fa30723;fa30723@lilly.com;john p patterson/am/lly;john p patterson/am/lly@lilly;patterson_john_p@lilly.com",
        primaryemail: "fa30723@lilly.com",
        uid: "fa30723"
        },
        {
                custodiantype: "internal",
                globalid: "2233814",
                displayname: "Darshan Doshi",
                primaryemail: "doshi_darshan@network.msg-q.lilly.com",
                mail: "c233814;c233814@lilly.com;doshi_darshan@network.lilly.com;doshi_darshan@network.msg-q.lilly.com",
                uid: "c233814"
        }
];

function compare(a, b) {
        if (a.mail.length > b.mail.length)
                return -1;
        if (a.mail.length < b.mail.length)
                return 1;
        return 0;
}

var ws_data_new = [];
var ws_content_data_new = [];
for (var i = 0; i < ws_data.length; i++) {
        ws_data_new[i] = ws_data[i].slice();
}

for (var i = 0; i < ws_content_data.length; i++) {
        var item = ws_content_data[i];
        for (var prop in item) {
                if (prop === 'mail') {
                        var str1 = item[prop];
                }
                if (prop === 'primaryemail') {
                        var str2 = item[prop];
                        var arr1 = str1.split(';');
                        var arr2 = str2.split(';');
                        var union = [...new Set([...arr1, ...arr2])];

                        var tempdata = '';
                        tempdata = JSON.parse('{"custodiantype" : "' + ws_content_data[i].custodiantype + '", \
                "displayname" : "' + ws_content_data[i].displayname + '", \
                "globalid" : "' + ws_content_data[i].globalid + '", \
                "uid" : "' + ws_content_data[i].uid + '", \
                "mail" : "' + (ws_content_data[i].mail !== null ? ws_content_data[i].mail : 'NA') + '"}');

                        ws_content_data_new.push(tempdata);
                        ws_content_data_new.sort(compare);
                }
        }
}

var flag = 0;
var limit = 1;
for (var item in ws_content_data_new) {
        var i = 0;
        var isTraversed = false;
        newitem = ws_content_data_new[item];
        for (var prop in newitem) {
                if (prop !== 'custodiantype') {

                        if (i < ws_data.length - 1) {
                                ws_data_new[i].push(newitem[prop]);
                                i += 2;
                        } else if (item[prop] !== 'NA') {
                                var str = newitem[prop];
                                var result = str.split(';');
                                var filteredResult = result.filter(function (el) {
                                        return el !== "";
                                });
                                console.log("filtered res", filteredResult);
                                filteredResult.forEach((newitem) => {
                                        if (flag < limit) {
                                                if (isTraversed === true) {
                                                        ws_data_new.push(['', newitem]);
                                                } else {
                                                        ws_data_new[i].push(newitem);
                                                        isTraversed = true;
                                                }
                                        } else if (newitem !== '' || newitem !== 'NA' || newitem !== null) {
                                                ws_data_new[i++].push(newitem);
                                        }
                                })
                                flag++;
                        }
                }
        }
};

var ws = XLSX.utils.aoa_to_sheet(ws_data_new);
var wscols = [{
        wch: 12
}];

for (var i = 0; i < ws_content_data_new.length; i++) {
        wscols.push({
                wch: 35
        })
}

var wsrows = [{
                hpt: 16
        },
        {
                hpx: 16
        },
];

ws['!rows'] = wsrows;
ws['!cols'] = wscols;

wb.Sheets["Download Vendor Sheet"] = ws;
var wbout = XLSX.write(wb, {
        bookType: 'xlsx',
        type: 'binary'
});

function printExcel(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
}

$("#button-a").click(function () {
        saveAs(new Blob([printExcel(wbout)], {
                        type: "application/octet-stream"
                }),
                'Download Vendor Sheet.xlsx');
});