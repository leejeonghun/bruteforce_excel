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

#ifndef BRUTEFORCE_EXCEL_DISPATCH_PARAMS_H_
#define BRUTEFORCE_EXCEL_DISPATCH_PARAMS_H_

#include <atlcomcli.h>

class dispatch_params final {
 public:
  dispatch_params();
  ~dispatch_params() = default;

  operator DISPPARAMS*() { return &params_; }
  bool set_filename(const wchar_t *filename);
  bool set_password(const wchar_t *password);

 private:
  VARIANT args_[5];
  VARIANT vt_readonly_;
  VARIANT vt_missing_;
  VARIANT &arg_filename_ = args_[_countof(args_) - 1];
  VARIANT &arg_readonly_ = args_[2];
  VARIANT &arg_password_ = args_[0];
  DISPPARAMS params_;
  CComBSTR filename_;
  CComVariant password_;
};

#endif  // BRUTEFORCE_EXCEL_DISPATCH_PARAMS_H_
