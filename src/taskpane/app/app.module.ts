import { NgModule, CUSTOM_ELEMENTS_SCHEMA } from "@angular/core";
import AppComponent from "./app.component";
import { BrowserAnimationsModule } from "@angular/platform-browser/animations";
import { BrowserModule } from "@angular/platform-browser";
import { MatFormFieldModule } from '@angular/material/form-field';
import {MatSelectModule} from '@angular/material/select';
import { MatInputModule } from '@angular/material/input';
import { MatMenuModule } from "@angular/material/menu";
import { MatButtonModule } from "@angular/material/button";
import { MatOptionModule, MatRippleModule } from "@angular/material/core";
import { MatRadioModule } from "@angular/material/radio";
import { MatTabsModule } from "@angular/material/tabs";
import { FormsModule, ReactiveFormsModule } from "@angular/forms";

const modules = [
  BrowserModule,
  BrowserAnimationsModule,
  MatInputModule,
  MatSelectModule,
  MatFormFieldModule,
  MatMenuModule,
  MatButtonModule,
  MatOptionModule,
  MatRadioModule,
  MatRippleModule,
  MatTabsModule,
  FormsModule,
  ReactiveFormsModule
];

@NgModule({
  imports: [...modules],
  exports: [...modules],
  declarations: [AppComponent],
  bootstrap: [AppComponent],
})
export default class AppModule {}
