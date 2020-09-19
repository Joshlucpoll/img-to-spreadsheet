<template>
  <div id="app">
    <div> 
      <input type="file" accept="image/*" @change="onFileChange($event)"/> 
      <button @click="onFileUpload()"> 
        Convert! 
      </button>
    </div>
  </div>
</template>

<script>
import Jimp from "jimp";
import Excel from "exceljs";

export default {
  name: 'App',
  methods: {

    toColumnName(num) {
      for (var ret = '', a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
        ret = String.fromCharCode(parseInt((num % b) / a) + 65) + ret;
      }
      return ret;
    },


    onFileChange(event) {
      this.file = event.target.files[0];
      this.fileName = this.file.name.split('.').slice(0, -1).join('.')
    },


    async downloadFile(workbook) {
      const buffer = await workbook.xlsx.writeBuffer()

      const blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      const url = window.URL.createObjectURL(blob);

      const download = document.createElement('a');
      download.setAttribute("href", url);
      download.setAttribute("download", `${this.fileName}-to-spreadsheet.xlsx`);

      download.style.display = 'none';
      document.body.appendChild(download);

      download.click();

      document.body.removeChild(download);
    },


    onFileUpload() {
      var reader = new FileReader();
      reader.onload = (e) => {

        Jimp.read(e.target.result)
        .then(image => {

          image.scaleToFit(200, 200);
          const pixelData = [];

          image.scan(0, 0, image.bitmap.width, image.bitmap.height, function(x, y, idx) {
          
            var red = this.bitmap.data[idx + 0];
            var green = this.bitmap.data[idx + 1];
            var blue = this.bitmap.data[idx + 2];

            pixelData.push([red, green, blue]);
          });

          const workbook = new Excel.Workbook();
          const sheet = workbook.addWorksheet('Image');

          let count = 0;
          
          for (let y = 0; y < image.bitmap.height; y++) {
            for (let x = 0; x < image.bitmap.width; x++) {
              for (let subPixel = 0; subPixel < 3; subPixel++) {

                const cellNum = this.toColumnName(x + 1) + (((y + 1) * 3) - 2 + subPixel).toString();
                const cell = sheet.getCell(cellNum);

                const data = pixelData[count][subPixel];
                const hexString = data.toString(16);
                let colour;
                
                if (subPixel === 0) {
                  colour = `00${hexString}0000`
                }
                if (subPixel === 1) {
                  colour = `0000${hexString}00`
                }
                if (subPixel === 2) {
                  colour = `000000${hexString}`
                }

                cell.value = data;

                cell.fill = {
                  type: "pattern",
                  pattern: "solid",
                  fgColor:{argb: colour}
                };

                cell.border = {
                  top: {style:'thin', color: {argb: colour}},
                  left: {style:'thin', color: {argb: colour}},
                  bottom: {style:'thin', color: {argb: colour}},
                  right: {style:'thin', color: {argb: colour}}
                };
              }
              count++;
            }
          }
          this.downloadFile(workbook);
        })
        .catch(err => {
          throw Error(err);
        })
      };

      reader.readAsArrayBuffer(this.file);
    },
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  padding-top: 60px;
}



/* Copyright (c) 2020 by Manuel Pinto (https://codepen.io/P1N2O/pen/pyBNzX)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. */
body {
  width: 100vw;
  height: 100vh;
  margin: 0;
  padding: 0;
  background: linear-gradient(-45deg, #ee7752, #e73c7e, #23a6d5, #23d5ab);
  background-size: 400% 400%;
  animation: gradient 15s ease infinite;
}

@keyframes gradient {
    0% {
        background-position: 0% 50%;
    }
    50% {
        background-position: 100% 50%;
    }
    100% {
        background-position: 0% 50%;
    }
}


</style>
