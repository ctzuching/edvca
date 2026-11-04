(function () {
  var form = document.getElementById("upload-form");
  var btn = document.getElementById("submit-btn");
  var err = document.getElementById("form-error");
  var MAX_MB = 10; // 與伺服器 MAX_UPLOAD_MB 預設值一致；實際以伺服器檢查為準

  function showError(msg) { err.textContent = msg; err.hidden = false; }

  form.addEventListener("submit", function (e) {
    err.hidden = true;
    var ids = ["old_file", "new_file"];
    for (var i = 0; i < ids.length; i++) {
      var f = document.getElementById(ids[i]).files[0];
      if (!f) { e.preventDefault(); showError("請選擇兩份 .docx 文件。"); return; }
      if (!/\.docx$/i.test(f.name)) { e.preventDefault(); showError("只接受 .docx 格式的 Word 文件。"); return; }
      if (f.size > MAX_MB * 1024 * 1024) { e.preventDefault(); showError("檔案超過 " + MAX_MB + " MB 上限。"); return; }
    }
    btn.disabled = true;
    btn.textContent = "分析中，請稍候…";
  });

  window.addEventListener("pageshow", function () {
    btn.disabled = false;
    btn.textContent = "開始分析";
  });
})();
