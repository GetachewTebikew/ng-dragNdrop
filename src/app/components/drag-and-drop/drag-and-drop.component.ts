import { Component, ElementRef, ViewChild } from '@angular/core';
import * as xlsx from 'xlsx';
const { read, utils: { sheet_to_json } } = xlsx;


@Component({
  selector: 'app-drag-and-drop',
  templateUrl: './drag-and-drop.component.html',
  styleUrls: ['./drag-and-drop.component.scss']
})
export class DragAndDropComponent {

  @ViewChild("fileDropRef", { static: false }) fileDropEl!: ElementRef;
  file?: File;
  data: any;
  accetableFormats = ['.csv', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel']

  /**
   * on file drop handler
   */
  onFileDropped($event: any) {
    this.prepareFilesList($event);
  }

  /**
   * handle file from browsing
   */
  fileBrowseHandler($event: any) {
    this.prepareFilesList($event.target.files);
  }


  /**
   * Convert Files list to normal array list
   * @param files (Files List)
   */
  prepareFilesList(files: Array<any>) {
    for (const file of files) {
      if (!this.accetableFormats.includes(file.type)) {
        alert("Please choose only excel files.");
        return;
      }
      file.progress = 0;
      if (file.size > 1024000) {
        alert(`The selected file is very large, please choose a file less than 1MB. T he size of this file is ${this.formatBytes(file.size)}`);
        return;
      }
      this.getData(file);
      this.file = file;
    }
    this.fileDropEl.nativeElement.value = "";
  }

  /**
   * format bytes
   * @param bytes (File size in bytes)
   * @param decimals (Decimals point)
   */
  formatBytes(bytes: number, decimals = 2) {
    if (bytes === 0) {
      return "0 Bytes";
    }
    const k = 1024;
    const dm = decimals <= 0 ? 0 : decimals;
    const sizes = ["Bytes", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + " " + sizes[i];
  }


  getData(file: File) {
    console.log(file.type);
    /* wire up file reader */
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: xlsx.WorkBook = xlsx.read(bstr, { type: 'binary' });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: xlsx.WorkSheet = wb.Sheets[wsname];

      /* save data */
      this.data = (xlsx.utils.sheet_to_html(ws));
      console.log(this.data);
    };
    reader.readAsBinaryString(file);
  }
}
