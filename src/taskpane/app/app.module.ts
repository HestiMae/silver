import { NgModule, CUSTOM_ELEMENTS_SCHEMA } from "@angular/core";
import { BrowserModule } from "@angular/platform-browser";
import { MatFormFieldModule } from '@angular/material/form-field';
import {MatSelectModule} from '@angular/material/select';
import { MatInputModule } from '@angular/material/input';
import AppComponent from "./app.component";
import { BrowserAnimationsModule } from "@angular/platform-browser/animations";

const modules = [
  BrowserModule,
  MatFormFieldModule,
  MatInputModule,
  MatSelectModule,
  BrowserAnimationsModule
];

@NgModule({
  declarations: [AppComponent],
  imports: [...modules],
  exports: [...modules],
  bootstrap: [AppComponent],
})
export default class AppModule {}
