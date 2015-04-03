// Copyright (c) 2015 jeonghun
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.

#include "bruteforce_excel/dispatch_params.h"

dispatch_params::dispatch_params() : params_({ args_, 0, _countof(args_), 0 }) {
  VariantInit(&vt_missing_);
  vt_missing_.vt = VT_ERROR;
  vt_missing_.lVal = DISP_E_PARAMNOTFOUND;

  VariantInit(&vt_readonly_);
  vt_readonly_.vt = VT_BOOL;
  vt_readonly_.boolVal = VARIANT_TRUE;

  for (auto& vt_arg : args_) {
    VariantInit(&vt_arg);
    vt_arg.vt = VT_VARIANT | VT_BYREF;
    vt_arg.pvarVal = &vt_missing_;
  }

  arg_readonly_.pvarVal = &vt_readonly_;
  arg_filename_.vt = VT_BSTR;
}

bool dispatch_params::set_filename(const wchar_t *filename) {
  filename_ = filename;
  arg_filename_.bstrVal = filename_;

  return filename != nullptr;
}

bool dispatch_params::set_password(const wchar_t *password) {
  password_ = password;
  arg_password_.pvarVal = &password_;

  return password != nullptr;
}
