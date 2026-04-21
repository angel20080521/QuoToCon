/* main.js – QuoToCon frontend enhancements */

(function () {
  'use strict';

  // ── File-picker label update ──────────────────────────────────────────────
  function setupFilePicker(inputId, labelId) {
    var input = document.getElementById(inputId);
    var label = document.getElementById(labelId);
    if (!input || !label) return;

    // Capture the original text so we can restore it when selection is cleared
    var originalText = label.textContent;

    input.addEventListener('change', function () {
      if (this.files && this.files.length > 0) {
        label.textContent = this.files[0].name;
        label.classList.add('has-file');
      } else {
        label.textContent = originalText;
        label.classList.remove('has-file');
      }
    });
  }

  setupFilePicker('quotation', 'lbl-quotation');
  setupFilePicker('template',  'lbl-template');

  // ── Drag-and-drop visual feedback ─────────────────────────────────────────
  document.querySelectorAll('.drop-area').forEach(function (area) {
    area.addEventListener('dragover', function (e) {
      e.preventDefault();
      this.classList.add('dragover');
    });
    area.addEventListener('dragleave', function () {
      this.classList.remove('dragover');
    });
    area.addEventListener('drop', function () {
      this.classList.remove('dragover');
    });
  });

  // ── Submit button loading state ───────────────────────────────────────────
  var form = document.querySelector('.upload-form');
  if (form) {
    form.addEventListener('submit', function () {
      var btn = document.getElementById('btn-submit');
      if (btn) {
        btn.textContent = '处理中，请稍候…';
        btn.disabled = true;
      }
    });
  }
})();
