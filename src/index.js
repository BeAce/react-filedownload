import XLSX from 'xlsx';
import FileSaver from 'file-saver';

function Workbook() {
  if (!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

class JSON2XLSX {
  datenum(v, date1904) {
    let date = v;
    if (date1904) date += 1462;
    const epoch = Date.parse(date);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
  }

  // create sheet array of array
  sheetFromArrayOfArrays(data) {
    const ws = {};
    const range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
    for (let R = 0; R !== data.length; ++R) {
      for (let C = 0; C !== data[R].length; ++C) {
        if (range.s.r > R) range.s.r = R;
        if (range.s.c > C) range.s.c = C;
        if (range.e.r < R) range.e.r = R;
        if (range.e.c < C) range.e.c = C;
        const cell = { v: data[R][C] };
        if (cell.v !== null) {
          const cellRef = XLSX.utils.encode_cell({ c: C, r: R });

          if (typeof cell.v === 'number') cell.t = 'n';
          else if (typeof cell.v === 'boolean') cell.t = 'b';
          else if (cell.v instanceof Date) {
            cell.t = 'n';
            cell.z = XLSX.SSF._table[14]; // eslint-disable-line no-underscore-dangle
            cell.v = this.datenum(cell.v);
          } else cell.t = 's';

          ws[cellRef] = cell;
        }
      }
    }
    if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
    return ws;
  }

  s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    // 返回低8位
    for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xff; // eslint-disable-line no-bitwise
    return buf;
  }

  sendFile(filename, wsName, titles, rows) {
    const data = [titles, ...rows];
    const wb = new Workbook();
    const ws = this.sheetFromArrayOfArrays(data);

    wb.SheetNames.push(wsName);
    wb.Sheets[wsName] = ws;
    const wbout = XLSX.write(wb, {
      bookType: 'xlsx',
      bookSST: true,
      type: 'binary',
    });

    FileSaver.saveAs(
      new Blob([this.s2ab(wbout)], { type: 'application/octet-stream' }),
      `${filename}.xlsx`
    );
  }
}

const json2xlsx = new JSON2XLSX();

export default json2xlsx;
