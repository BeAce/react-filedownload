import React from 'react';
import json2xlsx from '../src/index';
const titles = [
  "id",
  "name",
  "email",
];

const name = "文件导出";
const sheetName = "sheet1";
const rows = [];
const row = Array(titles.length)
  .fill()
  .map((_, i) => i + 1);
for (let i = 0; i < 10; i += 1) {
  rows.push(row);
}
class DownloadPage extends React.Component {
  download = () => {
    json2xlsx.sendFile(name, sheetName, titles, rows);
  }

  render() {
    return (
      <button onClick={this.download}>download</button>
    )
  }
}

export default DownloadPage;
