var win32 = ( function() {
  const ffi = require("ffi");
  const ref = require('ref');

  const voidPtr = ref.refType(ref.types.void);
  const stringPtr = ref.refType(ref.types.CString);

  const SWP_NOSIZE = 1;
  const SWP_NOMOVE = 2;
  const SW_RESTORE = 9;
  const HWND_NOTOPMOST = -2;
  const HWND_TOPMOST = -1;
  const HWND_TOP = 0;
  const SWP_SHOWWINDOW = 0x0040;

  const user32 = ffi.Library("user32.dll", {
    SetWindowPos: [
      "bool",
      ["long", "int32", "int32", "int32", "int32", "int32", "int32"]
    ],

    SetWindowTextA: ["bool", ["long", stringPtr]],

    SetForegroundWindow: ["bool", ["long"]],

    SetFocus: ["long", ["long"]],

    GetWindowTextA: ["long", ["long", stringPtr, "long"]],

    ShowWindow: ["bool", ["long", "int32"]],

    BringWindowToTop: ["bool", ["long"]],

    EnumWindows: ["bool", [voidPtr, "int32"]]
  });

  const api = {
    findWindow: function findWindow(s) {
      var res = null;

      const windowProc = ffi.Callback("bool", ["long", "int32"], function(hwnd) {
        const buf = new Buffer(255);

        user32.GetWindowTextA(hwnd, buf, 255);

        if (ref.readCString(buf, 0).includes(s)) {
          res = hwnd;

          return false;
        }

        return true;
      });

      user32.EnumWindows(windowProc, 0);

      return res;
    },

    setForegroundWindow: function setForegroundWindow(title) {
      const hwnd = api.findWindow(title);

      if (hwnd) {
        user32.ShowWindow(hwnd, SW_RESTORE);

        user32.SetWindowPos(
          hwnd,
          HWND_TOPMOST,
          0,
          0,
          0,
          0,
          SWP_NOMOVE | SWP_NOSIZE
        );

        user32.SetWindowPos(
          hwnd,
          HWND_NOTOPMOST,
          0,
          0,
          0,
          0,
          SWP_NOMOVE | SWP_NOSIZE
        );

        user32.SetFocus(hwnd);
      }
    },

    setWindowText: function setWindowText(hwnd, text) {
      hwnd && user32.SetWindowTextA(hwnd, new Buffer(text));
    },

    setFocus: function setFocus(hwnd) {
      hwnd && user32.SetFocus(hwnd);
    }
  };

  return api;
}());

module.exports = win32
