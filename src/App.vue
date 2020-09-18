<template>
  <div id="app">
    <div> 
      <input type="file" accept="image/*" @change="onFileChange($event)"/> 
      <button @click="onFileUpload()"> 
        Upload! 
      </button>
    </div>
    <a id="download">download</a>
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
    },

    async downloadFile(workbook) {
      const buffer = await workbook.xlsx.writeBuffer()
      console.log(buffer);
      let blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      let url = window.URL.createObjectURL(blob);

      document.getElementById("download").href = url;
      console.log(buffer);
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

          console.log(pixelData);

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


          console.log(workbook)


          this.downloadFile(workbook);

        })
        .catch(err => {
          throw Error(err);
        })
      };

      reader.readAsArrayBuffer(this.file);
    }
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
  margin-top: 60px;
}
</style>
