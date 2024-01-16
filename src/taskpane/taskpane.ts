/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import "zone.js"; // Required for Angular
import { platformBrowserDynamic } from "@angular/platform-browser-dynamic";
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import AppModule from "./app/app.module";
/* global console, document, Office */

Office.initialize = () => {
  // Bootstrap the app
  platformBrowserDynamic().bootstrapModule(AppModule).then(ref => {
    // Ensure Angular destroys itself on hot reloads.
    if (window['ngRef']) {
      window['ngRef'].destroy();
    }
    window['ngRef'] = ref;

    // Otherwise, log the boot error
  }).catch(err => console.error(err));
};
