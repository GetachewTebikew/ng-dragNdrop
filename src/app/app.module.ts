import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { AppComponent } from './app.component';
import { DragAndDropComponent } from './components/drag-and-drop/drag-and-drop.component';
import { DragDropDirective } from './directives/drag-drop.directive';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { MatProgressBarModule } from '@angular/material/progress-bar';
import { MatIconModule } from '@angular/material/icon';
import { MatButtonModule } from '@angular/material/button';

@NgModule({
  declarations: [
    AppComponent,
    DragAndDropComponent,
    DragDropDirective
  ],
  imports: [
    BrowserModule,
    BrowserAnimationsModule,

    // Angular Material modules
    MatProgressBarModule,
    MatIconModule,
    MatButtonModule,
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
