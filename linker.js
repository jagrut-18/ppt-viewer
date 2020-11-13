const { PythonShell } = require("python-shell");
const { ipcRenderer } = require("electron");
var fs = require("fs");
var child = require("child_process").execFile;

var loader = document.createElement("div");
loader.className = "loader";
var loaderBig = document.createElement("div");
loaderBig.className = "loader_big";

var overlay = document.createElement("div");
overlay.id = "overlay";
overlay.appendChild(loaderBig);

var mainDiv = document.getElementById("main_div");

var dropArea = document.getElementById("drop_area");
var inputFile = "";
var baseName = "";

var bodyHome = document.createElement("div");
bodyHome.id = "body_home";
bodyHome.className = "body_div";
bodyHome.tabIndex = "0";

var bodySave = document.createElement("div");
bodySave.id = "body_save";
bodySave.className = "body_div";

var bodyExit = document.createElement("div");
bodyExit.id = "body_exit";
bodyExit.className = "body_div";

bodyHome.style.width = (window.innerWidth - 215).toString() + "px";
bodyHome.style.height = (window.innerHeight - 120).toString() + "px";
bodySave.style.width = (window.innerWidth - 215).toString() + "px";
bodySave.style.height = (window.innerHeight - 110).toString() + "px";

var img = document.createElement("img");
img.id = "slide_img";
var rightDiv = document.getElementById("right_div");

var controlBar = document.createElement("div");
controlBar.id = "control_bar";
var next = document.createElement("img");
next.id = "next_button";
next.src = "assets/arrow.svg";
next.setAttribute("onclick", "control_next()");
var prev = document.createElement("img");
prev.id = "prev_button";
prev.src = "assets/arrow.svg";
prev.setAttribute("onclick", "control_prev()");
var fileName = document.createElement("input");
fileName.id = "file_name";
var currentPage = document.createElement("input");
currentPage.id = "current_page";
currentPage.type = "number";
currentPage.min = "1";
var totalPage = document.createElement("a");
totalPage.id = "total_page";
var totalPageDiv = document.createElement("div");
totalPageDiv.id = "pagecount_div";
totalPageDiv.appendChild(currentPage);
var divider = document.createElement("a");
divider.innerHTML = "/";
totalPageDiv.appendChild(divider);
totalPageDiv.appendChild(totalPage);
var divtocentercontrolbar = document.createElement("div");
divtocentercontrolbar.id = "div_to_center_controlbar";
divtocentercontrolbar.appendChild(prev);
divtocentercontrolbar.appendChild(totalPageDiv);
divtocentercontrolbar.appendChild(next);
controlBar.appendChild(fileName);
controlBar.appendChild(divtocentercontrolbar);

var page_count = 0;
var imgDir = "";

function loadPpt(message) {
  var data = message.split("==**==");
  if (data[0] != "error") {
    rightDiv.style.justifyContent = "start";
    img.src = data[0] + "\\Slide1.PNG";
    imgDir = data[0] + "/";
    rightDiv.innerHTML = "";
    rightDiv.appendChild(controlBar);
    bodyHome.appendChild(img);
    rightDiv.appendChild(bodyHome);
    bodyHome.style.display = "block";
    bodySave.style.display = "none";
    bodyExit.style.display = "none";
    rightDiv.appendChild(bodySave);
    rightDiv.appendChild(bodyExit);
    totalPage.innerHTML = data[1];
    currentPage.value = "1";
    cur_page = 1;
    currentPage.max = data[1];
    page_count = parseInt(data[1]);
    fileName.value = data[2];
    baseName = data[2];
    inputFile = data[3];
    bodyHome.focus();
  } else {
    bubble("Error opening file!", document.getElementById("menu_open"));
  }

  document.body.removeChild(overlay);
}

dropArea.addEventListener("drop", (e) => {
  e.preventDefault();
  e.stopPropagation();

  //   for (const f of e.dataTransfer.files) {
  //     console.log("File(s) you dragged here: ", f.path);
  //   }
  var dropFileExt = e.dataTransfer.files[0].path.split(".").pop();
  if (
    dropFileExt.toLowerCase() == "ppt" ||
    dropFileExt.toLowerCase() == "pptx"
  ) {
    document.body.appendChild(overlay);
    var options = {
      scriptPath: __dirname,
      args: ["drag", e.dataTransfer.files[0].path],
    };
    child("convert.exe", ["drag", e.dataTransfer.files[0].path], function (
      err,
      data
    ) {
      if (err) {
        console.error(err);
        return;
      }

      loadPpt(data.toString());
    });
    // var shell = new PythonShell("convert.exe", options);
    // shell.on("message", function (message) {
    //   loadPpt(message);
    // });
  } else {
    bubble(
      "Only ppt/pptx files are supported!",
      document.getElementById("menu_open")
    );
  }
});
dropArea.addEventListener("dragover", (e) => {
  dropArea.style.border = "4px solid #1389f0";
  dropArea.style.webkitBoxShadow = "0px 0px 89px -34px rgba(19,137,240,1";
  dropArea.style.mozBoxShadow = "0px 0px 89px -34px rgba(19,137,240,1";
  dropArea.style.boxShadow = "0px 0px 89px -34px rgba(19,137,240,1)";
  e.preventDefault();
  e.stopPropagation();
});
dropArea.ondragleave = (event) => {
  dropArea.style.border = "3px solid #85929e";

  dropArea.style.webkitBoxShadow = "";
  dropArea.style.mozBoxShadow = "";
  dropArea.style.boxShadow = "";
};

var cur_page = 1;
function control_next() {
  var img = document.getElementById("slide_img");
  if (cur_page >= page_count) {
    cur_page = 1;
  } else {
    cur_page++;
  }
  currentPage.innerHTML = cur_page.toString();
  img.src = imgDir + "Slide" + cur_page.toString() + ".PNG";
  currentPage.value = cur_page.toString();
}
function control_prev() {
  var img = document.getElementById("slide_img");
  if (cur_page <= 1) {
    cur_page = page_count;
  } else {
    cur_page--;
  }
  currentPage.innerHTML = cur_page.toString();
  img.src = imgDir + "Slide" + cur_page.toString() + ".PNG";
  currentPage.value = cur_page.toString();
}

bodyHome.onkeydown = function (e) {
  switch (e.keyCode) {
    case 37: //Left
      control_prev();
      break;
    case 38: //Up
      control_prev();
      break;
    case 39: //Right
      control_next();
      break;
    case 40: //Down
      control_next();
      break;
  }
};

currentPage.onkeydown = function (e) {
  if (e.keyCode == 13) {
    cur_page = parseInt(currentPage.value);
    img.src = imgDir + "Slide" + currentPage.value + ".PNG";
  }
};

window.addEventListener("resize", () => {
  bodyHome.style.width = (window.innerWidth - 215).toString() + "px";
  bodyHome.style.height = (window.innerHeight - 110).toString() + "px";
  bodySave.style.width = (window.innerWidth - 215).toString() + "px";
  bodySave.style.height = (window.innerHeight - 110).toString() + "px";
  bodyHome.style.margin = "auto";
});

document.getElementById("menu_home").className += " active";
document.getElementById("icon_home").style.fill = "#6490f1";

function select_menu(event, selected_id) {
  if (selected_id == "menu_save" && inputFile == "") {
    var menuOpen = document.getElementById("menu_open");
    bubble("Please open a file first", menuOpen);
    return true;
  }
  var menus = document.getElementsByClassName("menu");
  var icons = document.getElementsByClassName("menu_icon");
  var bodyDivs = document.getElementsByClassName("body_div");
  for (var i = 0; i < menus.length; i++) {
    menus[i].className = menus[i].className.replace(" active", "");
    icons[i].style.fill = "#85929e";
    bodyDivs[i].style.display = "none";
  }

  document.getElementById(selected_id).className += " active";
  document.getElementById(selected_id.replace("menu", "icon")).style.fill =
    "#6490f1";

  document.getElementById(selected_id.replace("menu", "body")).style.display =
    "block";
  if (selected_id == "menu_home") {
    divtocentercontrolbar.style.visibility = "visible";
  } else {
    divtocentercontrolbar.style.visibility = "hidden";
  }
  if (selected_id == "menu_exit") {
    ipcRenderer.send("close");
  }
}

function openFileDialog() {
  document.body.appendChild(overlay);
  var options = {
    scriptPath: __dirname,
    args: ["open"],
  };
  child("convert.exe", ["open"], function (err, data) {
    if (err) {
      console.error(err);
      return;
    }

    loadPpt(data.toString());
  });
  // var shell = new PythonShell("convert.exe", options);
  // shell.on("message", function (message) {
  //   loadPpt(message);
  // });
}

function resizable(el, factor) {
  var int = Number(factor) || 7.7;
  function resize() {
    el.style.width = (el.value.length + 1) * int + "px";
  }
  var e = "keyup,keypress,focus,blur,change".split(",");
  for (var i in e) el.addEventListener(e[i], resize, false);
  resize();
}
resizable(fileName, 10);

// function setImageSize() {
//   var originalHeight = img.naturalHeight;
//   var originalWidth = img.naturalWidth;
//   var setHeight = window.innerHeight - 120;
//   var setWidth = (img.offsetHeight * originalWidth) / originalHeight;
//   img.style.height = setHeight.toString() + "px";
//   img.style.width = setWidth.toString() + "px";
//   if (img.offsetWidth > window.innerWidth - 50 - leftDiv.offsetWidth) {
//     setWidth = window.innerWidth - 50 - leftDiv.offsetWidth;
//     setHeight = (img.offsetWidth * originalHeight) / originalWidth;
//     img.style.width = setWidth.toString() + "px";
//     img.style.height = setHeight.toString() + "px";
//   }
// }

var savePathDiv = document.createElement("div");
savePathDiv.id = "save_path_div";

var selectedPath = document.createElement("a");
selectedPath.id = "selected_path";
selectedPath.innerHTML = "Select folder to save file";

var getSavePathButton = document.createElement("button");
getSavePathButton.id = "get_save_path_button";
getSavePathButton.innerHTML = "Select";
getSavePathButton.addEventListener("click", () => {
  ipcRenderer.send("open-dialog");
});

var savePath = "";
ipcRenderer.on("got-folder", (event, path) => {
  if (path != null && path.slice(0, 9) != "E-r-r-o-r") {
    selectedPath.innerHTML = path;
    savePath = path;
  }
});

savePathDiv.appendChild(selectedPath);
savePathDiv.appendChild(getSavePathButton);

bodySave.appendChild(savePathDiv);

var convertTo = document.createElement("a");
convertTo.id = "convert_to";
convertTo.innerHTML = "Convert to";

var extensionsDiv = document.createElement("div");
extensionsDiv.id = "extensions_div";

var jpgDiv = document.createElement("div");
jpgDiv.id = "jpg_div";
jpgDiv.className = "ext_div";
var jpgRadio = document.createElement("div");
jpgRadio.id = "jpg_radio";
jpgRadio.className = "ext_radio";
var jpgHeader = document.createElement("a");
jpgHeader.className = "save_header";
jpgHeader.innerHTML = "JPG";
var pdfDiv = document.createElement("div");
pdfDiv.id = "pdf_div";
pdfDiv.className = "ext_div";
var pdfRadio = document.createElement("div");
pdfRadio.id = "pdf_radio";
pdfRadio.className = "ext_radio";
var pdfHeader = document.createElement("a");
pdfHeader.className = "save_header";
pdfHeader.innerHTML = "PDF";
pdfDiv.style.border = "2px solid #fe346e";
pdfRadio.style.backgroundColor = "#fe346e";
var pngDiv = document.createElement("div");
pngDiv.id = "png_div";
pngDiv.className = "ext_div";
var pngRadio = document.createElement("div");
pngRadio.id = "png_radio";
pngRadio.className = "ext_radio";
var pngHeader = document.createElement("a");
pngHeader.className = "save_header";
pngHeader.innerHTML = "PNG";

jpgDiv.appendChild(jpgRadio);
jpgDiv.appendChild(jpgHeader);
jpgDiv.addEventListener("click", (event) => {
  selectExt("jpg_div");
});
pdfDiv.appendChild(pdfRadio);
pdfDiv.appendChild(pdfHeader);
pdfDiv.addEventListener("click", (event) => {
  selectExt("pdf_div");
});
pngDiv.appendChild(pngRadio);
pngDiv.appendChild(pngHeader);
pngDiv.addEventListener("click", (event) => {
  selectExt("png_div");
});

bodySave.appendChild(convertTo);

extensionsDiv.appendChild(pdfDiv);
extensionsDiv.appendChild(pngDiv);
extensionsDiv.appendChild(jpgDiv);

var convertButton = document.createElement("button");
convertButton.id = "convert_button";
convertButton.innerHTML = "Convert and Save";

bodySave.appendChild(extensionsDiv);
bodySave.appendChild(convertButton);

var selected_ext = "pdf";

function selectExt(selected_id) {
  var exts = document.getElementsByClassName("ext_div");
  var radios = document.getElementsByClassName("ext_radio");
  for (var i = 0; i < exts.length; i++) {
    exts[i].style.border = "2px solid #d1d1d19c";
    radios[i].style.backgroundColor = "#d1d1d19c";
  }

  document.getElementById(selected_id).style.border = "2px solid #fe346e";
  document.getElementById(
    selected_id.replace("div", "radio")
  ).style.backgroundColor = "#fe346e";

  if (selected_id == "pdf_div") {
    selected_ext = "pdf";
  } else if (selected_id == "jpg_div") {
    selected_ext = "jpg";
  } else {
    selected_ext = "png";
  }
}

convertButton.addEventListener("click", (event) => {
  if (selected_ext != "" && savePath != "") {
    convertButton.appendChild(loader);
    convertButton.disabled = true;

    child(
      "convert.exe",
      ["save", savePath, inputFile.toString(), selected_ext],
      function (err, data) {
        if (err) {
          console.error(err);
          return;
        }
        message = data.toString();
        message = message.replace("\n", "").toString();
        message = message.replace("\r", "").toString();
        if (message == "success") {
          convertButton.removeChild(loader);
          convertButton.disabled = false;
          bubble("File converted successfully at " + savePath, convertButton);
        } else {
          bubble("Error in conversion", convertButton);
          convertButton.removeChild(loader);
          convertButton.disabled = false;
        }

        selectedPath.innerHTML = "Select folder to save file";
        savePath = "";
      }
    );
  } else if (savePath == "") {
    bubble("Please select folder!", getSavePathButton);
  } else {
    bubble("Please select extension!", extensionsDiv);
  }
});

fileName.onkeydown = function (e) {
  if (e.keyCode == 13) {
    fs.rename(inputFile, inputFile.replace(baseName, fileName.value), function (
      err
    ) {
      if (err == null) {
        inputFile = inputFile.replace(baseName, fileName.value);
        bubble("File rename successful", fileName);
      } else {
      }
    });
  }
};

var bubbleDiv = document.createElement("div");
bubbleDiv.id = "bubble_div";
bubbleDiv.style.display = "none";
var bubbleText = document.createElement("p");
bubbleText.id = "bubble_text";
bubbleText.innerHTML = "Hello";

bubbleDiv.appendChild(bubbleText);
mainDiv.appendChild(bubbleDiv);

function bubble(message, element) {
  element.parentNode.appendChild(bubbleDiv);
  bubbleText.innerHTML = message;
  var rect = element.getBoundingClientRect();
  bubbleDiv.style.display = "block";
  bubbleDiv.style.left = rect.left.toString() + "px";
  bubbleDiv.style.top = rect.bottom.toString() + "px";
  setTimeout(function () {
    bubbleDiv.style.display = "none";
  }, 5000);
}

document.getElementById("win_close").addEventListener("click", (event) => {
  ipcRenderer.send("close");
});

var winMax = true;

document.getElementById("win_max").addEventListener("click", (event) => {
  if (winMax) {
    ipcRenderer.send("maximize");
  } else {
    ipcRenderer.send("original");
  }
  winMax = !winMax;
});
document.getElementById("win_min").addEventListener("click", (event) => {
  ipcRenderer.send("minimize");
});
