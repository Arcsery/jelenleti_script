var xl = require('excel4node');
const fs = require('fs');
const { format } = require('date-fns');

var wb = new xl.Workbook()
var ws = wb.addWorksheet('Sheet 1');

const year = 2023
const month = 11

function getDaysInMonth(month) {
    return new Date(0, month, 0).getDate();
}

var emptyGreyCell = wb.createStyle({
    fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: '#808080',
        fgColor: '#808080',
    },
    alignment: {
        wrapText: true,
        horizontal: 'center',
        vertical: 'center'
    },
    border: {
        left: {
            style: "thin"
        },
        right: {
            style: "thin"
        },
        top: {
            style: "thin"
        },
        bottom: {
            style: "thin"
        }
    },
})


var emptyWhiteCell = wb.createStyle({
    fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: '#FFFFFF',
        fgColor: '#FFFFFF',
    },
    alignment: {
        wrapText: true,
        horizontal: 'center',
        vertical: 'center'
    },
    border: {
        left: {
            style: "thin"
        },
        right: {
            style: "thin"
        },
        top: {
            style: "thin"
        },
        bottom: {
            style: "thin"
        }
    },
})

var filledDateCell = wb.createStyle({
    font: {
        size: 12,
      },
      fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: '#FFFFFF',
        fgColor: '#FFFFFF',
    },
    alignment: {
        wrapText: true,
        horizontal: 'center',
        vertical: 'center'
    },
    border: {
        left: {
            style: "thin"
        },
        right: {
            style: "thin"
        },
        top: {
            style: "thin"
        },
        bottom: {
            style: "thin"
        }
    },
    numberFormat: 'HH:MM:SS'
})

var headerStyle = wb.createStyle({
    font: {
        size: 10,
        style: 'bold'
      },    
    border: {
        left: {
            style: "thin"
        },
        right: {
            style: "thin"
        },
        top: {
            style: "thin"
        },
        bottom: {
            style: "thin"
        }
    },
    alignment: {
        wrapText: true,
        horizontal: 'center',
        vertical: 'center'
      },
  });

  var dateStyle = wb.createStyle({
    border: {
        left: {
            style: "thin"
        },
        right: {
            style: "thin"
        },
        top: {
            style: "thin"
        },
        bottom: {
            style: "thin"
        }
    },
    alignment: {
        wrapText: true,
        horizontal: 'center',
        vertical: 'center'
      },
    numberFormat: 'yyyy-mm-dd'
})

function addLeadingZero(time) {
    const [hours, minutes] = time.split(":");
    const formattedHours = hours.length === 1 ? `0${hours}` : hours;
    const formattedMinutes = minutes.length === 1 ? `0${minutes}` : minutes;
    return `${formattedHours}:${formattedMinutes}`;
}

var formattedStart = ""
var formattedEnd = ""

//File beolvasás
fs.readFile("data.txt", 'utf8', function (err, data) {
    if (err) throw err;

    //object-be rendezés, day formattedStart formattedEnd
    var dataArray = data.split("\n")
    .filter(line => line.trim() !== "")
    .map(line => {
        var [day, start, end] = line.split(",");
        day = parseInt(day);
        if(start != undefined && end != undefined){
            formattedStart = format(new Date(`2000-01-01T${addLeadingZero(start)}`), 'HH:mm:ss');
            formattedEnd = format(new Date(`2000-01-01T${addLeadingZero(end)}`), 'HH:mm:ss');
        }
        return { day, formattedStart, formattedEnd };
    });
    //A D E F oszlop szélesség
    ws.column(1).setWidth(10);
    ws.column(4).setWidth(15);
    ws.column(5).setWidth(20);
    ws.column(6).setWidth(18);

    //Második sor magasság
    ws.row(2).setHeight(30);

    //Fejlécek
    ws.cell(1, 1)
        .string('Buza Benjámin - jelenléti ív gyakornoki programban');
    ws.cell(2, 1)
        .string('dátum')
        .style(headerStyle);    
    ws.cell(2, 2)
        .string('érkezés')
        .style(headerStyle);
    ws.cell(2, 3)
        .string('távozás')
        .style(headerStyle);
    ws.cell(2, 4)
        .string('óra összesen')
        .style(headerStyle);
    ws.cell(2, 5)
        .string('Téma')
        .style(headerStyle);
    ws.cell(2, 6)
        .string('szakmai gyakorlat vezető aláírása  ')
        .style(headerStyle);    

    //A oszlop feltöltése a napokkal
    var days = getDaysInMonth(month)
    for (let i = 1; i <= days; i++) {
        const datum = new Date(year, month-1, i+1);
        ws.cell(2+i,1)
        .date(datum)
        .style(dateStyle)

    //minden oszlopot szürkére szinezek

        ws.cell(2+i,2)
        .style(emptyGreyCell)
        ws.cell(2+i,3)
        .style(emptyGreyCell) 
        ws.cell(2+i,4)
        .style(emptyGreyCell) 
        ws.cell(2+i,5)
        .style(emptyGreyCell) 
        ws.cell(2+i,6)
        .style(emptyGreyCell)    
    }

    //érkezés távozás feltöltése és a velük egy sorban lévő cellák szinezése
    for (let i = 0; i < dataArray.length; i++) {
        formattedStart = dataArray[i].formattedStart
        firmattedEnd = dataArray[i].formattedEnd
        ws.cell(dataArray[i].day + 2, 2)
        .string(formattedStart)
        .style(filledDateCell)

        ws.cell(dataArray[i].day + 2, 3)
        .string(formattedEnd)
        .style(filledDateCell)

        ws.cell(dataArray[i].day + 2, 4)
        .formula(`=C${dataArray[i].day + 1}-B${dataArray[i].day + 1}`)
        .style(filledDateCell)

        ws.cell(dataArray[i].day + 2, 5)
        .style(emptyWhiteCell)

        ws.cell(dataArray[i].day + 2, 6)
        .style(emptyWhiteCell)
    }

    wb.write('Excel.xlsx');
});

getDaysInMonth(11)

