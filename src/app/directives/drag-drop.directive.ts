import { Directive, EventEmitter, HostBinding, HostListener, Output } from '@angular/core';

@Directive({
  selector: '[appDragDrop]'
})
export class DragDropDirective {

  @HostBinding('class.fileover') fileOver!: boolean;
  @Output() fileDropped = new EventEmitter();

  // listen for dragover
  @HostListener('dragover', ['$event']) onDragOver(evt:any) {
    evt.preventDefault();
    evt.stopPropagation();
    this.fileOver = true;
  }

  // listener for dragleave
  @HostListener('dragleave', ['$event']) public onDragLeave(evt: any) {
    evt.preventDefault();
    evt.stopPropagation();
    this.fileOver = false;
  }

  // listen for drop
  @HostListener('drop', ['$event']) public ondrop(evt: any) {
    evt.preventDefault();
    evt.stopPropagation();
    this.fileOver = false;
    let files = evt.dataTransfer.files;
    if (files.length > 0) {
      this.fileDropped.emit(files);
    }
  }

}
