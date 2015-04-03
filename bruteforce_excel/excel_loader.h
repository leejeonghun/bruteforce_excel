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

#ifndef BRUTEFORCE_EXCEL_BRUTEFORCE_EXCEL_EXCEL_LOADER_H_
#define BRUTEFORCE_EXCEL_BRUTEFORCE_EXCEL_EXCEL_LOADER_H_

#include <atlcomcli.h>
#include "bruteforce_excel/dispatch_params.h"

class excel_loader final {
 public:
  excel_loader();
  ~excel_loader();

  bool init();
  void uninit();
  bool set_file(const wchar_t *file);
  bool try_open(const wchar_t *passwd);

 private:
  enum {
    DISPID_APP_WORKBOOKS = 0,
    DISPID_APP_QUIT,
    DISPID_WBS_OPEN,
    DISPID_WBS_CLOSE,
    COUNT_OF_DISPID,
  };

  inline HRESULT create_excel_instance(IDispatch** disp_pp) const;
  inline DISPID get_dispid(IDispatch* disp_ptr, const wchar_t *name) const;

  bool init_ = false;
  CComPtr<IDispatch> app_disp_ptr_;
  CComPtr<IDispatch> wbs_disp_ptr_;
  dispatch_params open_params_;
  DISPID id_[COUNT_OF_DISPID];
};

#endif  // BRUTEFORCE_EXCEL_BRUTEFORCE_EXCEL_EXCEL_LOADER_H_
