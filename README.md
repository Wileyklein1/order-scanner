# Order Scanner

An offline web app for iPad: scan product barcodes (camera or Bluetooth scanner) into an order built from an order-form spreadsheet you load on the device, then export the order as an .xlsx.

- No data is stored in this app or repository. The order form you load, and every order you take, stay on your device (IndexedDB).
- Works with no signal after the first visit. Add it to the Home Screen and launch it from there.
- Libraries are bundled in `vendor/`: [SheetJS](https://sheetjs.com) (Apache-2.0) and [zxing-wasm](https://github.com/Sec-ant/zxing-wasm) (MIT; zxing-cpp, Apache-2.0).
