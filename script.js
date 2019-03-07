var wb = XLSX.utils.book_new();

wb.SheetNames.push("Download Vendor Sheet");

var ws_data = [
        ['NAME'],
        [],
        ['GID'],
        [],
        ['UID'],
        [],
        ['EMAIL ALIAS'],
        [''],
        ['']
];



var ws_data_new = [];
var ws_content_data_new = [] ;

for (var i = 0; i < ws_data.length; i++)
        ws_data_new[i] = ws_data[i].slice();

var ws_content_data = [
        {
                custodiantype: "external",
                displayname: "John E Peterson",
                globalid: "10003",
                mail: null,
                notesemail: "notesemail1@lilly.com",
                primaryemail: "primaryemail1@lilly.com",
                uid: "mail"
        },
        {
                custodiantype: "externa2l",
                displayname: "John E Peterson",
                globalid: "10003",
                mail: "mail2@lilly.com",
                notesemail: null,
                primaryemail: "primaryemail2@lilly.com",
                uid: "notesemail"
        },
        {
                custodiantype: "exter235nal",
                displayname: "John E Peterson",
                globalid: "10003",
                mail: "mail3@lilly.com",
                notesemail: "notesemail3@lilly.com",
                primaryemail: null,
                uid: "primaryemail"
        }
];


ws_content_data.forEach(item => {
       var tempdata = JSON.parse(`{
        "custodiantype": "${item.custodiantype}",
        "displayname": "${item.displayname}",
        "globalid": "${item.globalid}",
        "uid": "${item.uid}",
        "mail": "${item.mail !== null ? item.mail:'NA'}",
        "notesemail": "${item.notesemail !== null ? item.notesemail:'NA'}",
        "primaryemail": "${item.primaryemail !== null ? item.primaryemail:'NA'}"
       }`)
       console.log(tempdata['mail']);
       ws_content_data_new.push(tempdata) 
}); 

ws_content_data_new.forEach(item => {
        var i = 0;
        for (var prop in item) {
                if(prop !== 'custodiantype'){
                        if (item[prop] !== null && item[prop] === toString(item[prop]) && item[prop].indexOf(';') > -1) {
                                item[prop] = item[prop].split(';').join(',');
                        }
        
                        if (i < ws_data.length - 3) {
                                ws_data_new[i].push(item[prop]);
                                i += 2;
                        }
                        else {
                                if(item[prop] !== 'NA'){
                                        ws_data_new[i].push(item[prop]);
                                        i += 1;
                                }                                                             
                        }
                }               
        }
});

var ws = XLSX.utils.aoa_to_sheet(ws_data_new);

var wscols = [
        { wch: 12 },
        { wch: 40 },
        { wch: 40 },
        { wch: 40 }
];
var wsrows = [
        { hpt: 12 }, 
        { hpx: 16 },
];

ws['!rows'] = wsrows;
ws['!cols'] = wscols;

wb.Sheets["Download Vendor Sheet"] = ws;
var wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
function printExcel(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
}

$("#button-a").click(function () {
        saveAs(new Blob([printExcel(wbout)], { type: "application/octet-stream" }),
                'Download Vendor Sheet.xlsx');
});
