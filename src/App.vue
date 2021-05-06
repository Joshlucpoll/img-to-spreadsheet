<template>
  <div id="app">
    <div class="title">Image to Spreadsheet Converter</div>
    <div class="col">
      <div id="file-input-container">
        <input
          id="file"
          type="file"
          accept="image/*"
          @change="onFileChange($event)"
        />
        <label class="file-label button" for="file">
          <svg
            width="1em"
            height="1em"
            viewBox="0 0 16 16"
            id="file-icon"
            class="bi bi-file-earmark-image"
            fill="currentColor"
            xmlns="http://www.w3.org/2000/svg"
          >
            <path
              fill-rule="evenodd"
              d="M12 16a2 2 0 0 0 2-2V4.5L9.5 0H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h8zM3 2a1 1 0 0 1 1-1h5.5v2A1.5 1.5 0 0 0 11 4.5h2V10l-2.083-2.083a.5.5 0 0 0-.76.063L8 11 5.835 9.7a.5.5 0 0 0-.611.076L3 12V2z"
            />
            <path
              fill-rule="evenodd"
              d="M6.502 7a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3z"
            />
          </svg>
          {{ fileText }}
        </label>
        <div
          class="button convert"
          v-if="fileText !== 'Feed me images...'"
          @click="onFileUpload()"
        >
          Convert!
          <div v-if="converting" class="loader"></div>
        </div>
      </div>
      <div class="button subpixel-sw" @click="subPixelSwitch()">
        <span class="material-icons">{{
          checked ? "check_box" : "check_box_outline_blank"
        }}</span>
        Include Subpixels
      </div>
      <!-- <input type="checkbox" class="toggle-switch oversize" id="toggle3">
        <label for="toggle3">Oversized knob</label> -->
      <!-- <input type="checkbox" id="checkbox" class="toggle-switch oversize" v-model="checked">
        <label for="checkbox">Include Subpixels</label> -->
    </div>
    <div class="credit">
      Made by
      <a
        href="https://joshlucpoll.com"
        target="_blank"
        rel="noopener noreferrer"
        >Joshlucpoll</a
      >
    </div>
  </div>
</template>

<script>
import Jimp from "jimp";
import Excel from "exceljs";

export default {
  name: "App",

  data() {
    return {
      fileText: "Feed me images...",
      converting: false,
      workbookURL: { whole: false, sub: false },
      checked: false,
    };
  },
  methods: {
    toColumnName(num) {
      for (var ret = "", a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
        ret = String.fromCharCode(parseInt((num % b) / a) + 65) + ret;
      }
      return ret;
    },

    onFileChange(event) {
      this.file = event.target.files[0];
      this.fileText = this.file.name;
      this.fileName = this.file.name
        .split(".")
        .slice(0, -1)
        .join(".");
      this.workbookURL = { whole: false, sub: false };

      if (this.file) {
        this.displayPreview();
      }
    },

    subPixelSwitch() {
      this.checked = !this.checked;
    },

    async downloadFile(url) {
      const download = document.createElement("a");
      download.setAttribute("href", url);
      download.setAttribute("download", `${this.fileName}-to-spreadsheet.xlsx`);

      download.style.display = "none";
      document.body.appendChild(download);

      download.click();

      document.body.removeChild(download);
    },

    sleep(ms) {
      return new Promise((resolve) => setTimeout(resolve, ms));
    },

    convert() {
      var reader = new FileReader();
      reader.onload = (e) => {
        this.converting = true;

        Jimp.read(e.target.result)
          .then(async (image) => {
            image.scaleToFit(400, this.checked ? 200 : 600);
            const pixelData = [];

            image.scan(0, 0, image.bitmap.width, image.bitmap.height, function(
              x,
              y,
              idx
            ) {
              var red = this.bitmap.data[idx + 0];
              var green = this.bitmap.data[idx + 1];
              var blue = this.bitmap.data[idx + 2];

              pixelData.push([red, green, blue]);
            });

            const workbook = new Excel.Workbook();
            const sheet = workbook.addWorksheet("Image", {
              properties: { defaultColWidth: this.checked ? 10 : 3 },
            });

            const componentToHex = (c) => {
              var hex = c.toString(16);
              return hex.length == 1 ? "0" + hex : hex;
            };

            let count = 0;

            if (this.checked) {
              for (let row = 0; row < image.bitmap.height; row++) {
                for (let col = 0; col < image.bitmap.width; col++) {
                  for (let subpixel = 0; subpixel < 3; subpixel++) {
                    const cellNum =
                      this.toColumnName(col + 1) +
                      ((row + 1) * 3 - 2 + subpixel);
                    const cell = sheet.getCell(cellNum);

                    let colour = "00";
                    if (subpixel == 0) {
                      colour += componentToHex(pixelData[count][0]) + "0000";
                    }
                    if (subpixel == 1) {
                      colour +=
                        "00" + componentToHex(pixelData[count][1]) + "00";
                    }
                    if (subpixel == 2) {
                      colour += "0000" + componentToHex(pixelData[count][2]);
                    }

                    cell.fill = {
                      type: "pattern",
                      pattern: "solid",
                      fgColor: { argb: colour },
                    };
                  }
                  count++;
                }
              }
            } else {
              for (let y = 0; y < image.bitmap.height; y++) {
                for (let x = 0; x < image.bitmap.width; x++) {
                  const cellNum = this.toColumnName(x + 1) + (y + 1).toString();
                  const cell = sheet.getCell(cellNum);

                  let colour =
                    "00" +
                    componentToHex(pixelData[count][0]) +
                    componentToHex(pixelData[count][1]) +
                    componentToHex(pixelData[count][2]);

                  cell.fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: { argb: colour },
                  };

                  count++;
                }
                await this.sleep(1);
              }
            }

            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], {
              type:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });
            const url = window.URL.createObjectURL(blob);

            if (this.checked) {
              this.workbookURL.sub = url
            }
            else {
              this.workbookURL.whole = url
            }
            this.converting = false;

            this.downloadFile(url);
          })
          .catch((err) => {
            throw Error(err);
          });
      };

      reader.readAsArrayBuffer(this.file);
    },

    onFileUpload() {
      if (!this.checked && this.workbookURL.whole) {
        this.downloadFile(this.workbookURL.whole);
      } else if (this.checked && this.workbookURL.sub) {
        this.downloadFile(this.workbookURL.sub);
      } else {
        this.convert();
      }
    },

    displayPreview() {
      var reader = new FileReader();

      reader.onload = (e) => {
        const img = document.body.contains(document.getElementById("preview"))
          ? document.getElementById("preview")
          : document.createElement("img");

        img.id = "preview";
        img.setAttribute("src", e.target.result);

        document.getElementById("file-input-container").before(img);
      };

      reader.readAsDataURL(this.file);
    },
  },
};
</script>

<style>
body {
  font-family: "Kumbh Sans", sans-serif;
  width: 100%;
  height: 100%;
  margin: 0;
  padding: 0;
  position: fixed;
  overflow: hidden;
}

#app {
  padding: 10vh 0;
  height: calc(100% - 20vh);
  width: 100%;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  overflow: hidden;

  display: flex;
  align-items: center;
  justify-content: space-around;
  flex-direction: column;
}

.title {
  color: white;
  font-weight: bold;
  font-size: 40px;
}

.col {
  /* height: 20vh; */
  display: flex;
  flex-direction: column;
  justify-content: space-around;
  align-items: center;
}

#file {
  width: 0.1px;
  height: 0.1px;
  opacity: 0;
  overflow: hidden;
  position: absolute;
  z-index: -1;
}

#file-input-container {
  display: flex;
  align-items: center;
  justify-content: center;
}

.file-label {
  margin: 0 20px;
}

.button {
  padding: 15px;
  padding-bottom: 10px;
  font-size: 22px;
  /* text-align: center; */
  text-decoration: none;
  color: white;
  text-shadow: 0 -1px -1px #0f864a;
  cursor: pointer;

  background-color: #313d53;
  border-radius: 4px;
  box-shadow: 0 4px 0 rgba(31, 39, 53, 0.75), 0 5px 5px 1px rgba(0, 0, 0, 0.4);

  transition: all 0.15s ease-in-out;
  user-select: none;

  /* display: flex;
  align-items: stretch; */
}

.button:hover {
  text-shadow: 0 -1px -1px #119d57;
  background-color: #273349;
  box-shadow: 0 4px 0 rgba(31, 39, 53, 1), 0 5px 5px 1px rgba(0, 0, 0, 0.4);
}

.button:active {
  margin-bottom: -4px;
  box-shadow: none;
}

.material-icons {
  /* margin-right: 5px; */
  transform: translateY(3px);
}

#file-icon {
  height: 18px;
  margin-right: 5px;
  transform: translateY(1px);
}

#preview {
  margin-bottom: 3vw;
  height: 20vh;
  position: relative;
  border: black solid 2px;
  border-radius: 2px;
  box-shadow: 0 4px 5px rgba(0, 0, 0, 0.226), 0 5px 5px 3px rgba(0, 0, 0, 0.226);
}

.convert {
  display: inline;
}

.loader {
  display: inline-block;
  border: 1px solid #f3f3f3;
  border-radius: 50%;
  border-top: 1px solid #3498db;
  width: 15px;
  height: 15px;
  -webkit-animation: spin 2s linear infinite; /* Safari */
  animation: spin 2s linear infinite;
}

.subpixel-sw {
  margin-top: 5vw;
}

.credit {
  background-color: #313d53;
  color: rgba(256, 256, 256, 0.4);
  padding: 10px;
  position: absolute;
  border-top-right-radius: 10px;
  bottom: 0;
  left: 0;
}

.credit a {
  color: white;
  text-decoration: none;
}

.credit a::before {
  content: "";
  position: absolute;
  width: 100%;
  height: 2px;
  bottom: 0;
  left: 0;
  background-color: yellow;
  visibility: hidden;
  -webkit-transform: scaleX(0);
  transform: scaleX(0);
  -webkit-transition: all 0.3s ease-in-out 0s;
  transition: all 0.3s ease-in-out 0s;
}

.credit a:hover {
  color: yellow;
}

.credit a:hover::before {
  visibility: visible;
  -webkit-transform: scaleX(1);
  transform: scaleX(1);
}

@-webkit-keyframes spin {
  0% {
    -webkit-transform: rotate(0deg);
  }
  100% {
    -webkit-transform: rotate(360deg);
  }
}

@keyframes spin {
  0% {
    transform: rotate(0deg);
  }
  100% {
    transform: rotate(360deg);
  }
}

/* Copyright (c) 2020 by Manuel Pinto (https://codepen.io/P1N2O/pen/pyBNzX)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. */
#app {
  background: linear-gradient(-45deg, #245501, #538d22, #73a942);
  background-size: 400% 400%;
  animation: gradient 15s ease infinite;
}

@keyframes gradient {
  0% {
    background-position: 0% 0%;
  }
  50% {
    background-position: 100% 100%;
  }
  100% {
    background-position: 0% 0%;
  }
}
</style>
