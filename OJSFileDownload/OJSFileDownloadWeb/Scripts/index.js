/// <reference path="../App.js" />

(function () {
    "use strict";

    // 新しいページが読み込まれるたびに初期化関数を実行する必要があります
    //Office.initialize = function (reason)
    {
        $(document).ready(function () {
            app.initialize();

            $('#pdf_XmlHttpRequest').on('click', { file: 'sample.pdf' }, xmlHttpRequestDownload);
            $('#png_XmlHttpRequest').on('click', { file: 'sample.png' }, xmlHttpRequestDownload);
        });
    };

    function xmlHttpRequestDownload(event) {
        var fileName = event.data.file;
        var oReq = new XMLHttpRequest();
        oReq.open("GET", fileName, true);
        oReq.responseType = "blob";

        oReq.onload = function (oEvent) {
            var blob = oReq.response; // Note: not oReq.responseText
            if (blob) {
                saveFile(blob, fileName);
            }
        };

        oReq.send(null);
    }

    function saveFile(blob, fileName) {
        if (window.navigator.msSaveOrOpenBlob) {
            window.navigator.msSaveOrOpenBlob(blob, fileName);
        } else {
            var link = document.createElement('a');
            link.href = window.URL.createObjectURL(blob);
            link.download = fileName;
            link.click();
        }
    }
})();