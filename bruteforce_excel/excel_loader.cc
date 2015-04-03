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

#include "bruteforce_excel/excel_loader.h"
#include <cstdint>

excel_loader::excel_loader() {
  for (auto &id : id_) id = -1;
  CoInitializeEx(0, COINIT_MULTITHREADED);
}

excel_loader::~excel_loader() {
  uninit();
  CoUninitialize();
}

bool excel_loader::init() {
  if (init_ == false) {
    CComPtr<IDispatch> app_disp_ptr;
    HRESULT hr = create_excel_instance(&app_disp_ptr);
    if (SUCCEEDED(hr) && app_disp_ptr != nullptr) {
      id_[DISPID_APP_WORKBOOKS] = get_dispid(app_disp_ptr, L"Workbooks");
      id_[DISPID_APP_QUIT] = get_dispid(app_disp_ptr, L"Quit");
    }

    CComPtr<IDispatch> wbs_disp_ptr;
    if (id_[DISPID_APP_WORKBOOKS] && id_[DISPID_APP_QUIT]) {
      DISPPARAMS params = { 0 };
      VARIANT result = { 0 };
      VariantInit(&result);
      hr = app_disp_ptr->Invoke(id_[DISPID_APP_WORKBOOKS],
        IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params,
        &result, nullptr, nullptr);
      if (SUCCEEDED(hr) && result.pdispVal != nullptr) {
        wbs_disp_ptr = result.pdispVal;
        result.pdispVal->Release();
      }
    }

    if (wbs_disp_ptr != nullptr) {
      id_[DISPID_WBS_OPEN] = get_dispid(wbs_disp_ptr, L"Open");
      id_[DISPID_WBS_CLOSE] = get_dispid(wbs_disp_ptr, L"Close");
    }

    if (id_[DISPID_WBS_OPEN] && id_[DISPID_WBS_CLOSE]) {
      wbs_disp_ptr_ = wbs_disp_ptr;
      app_disp_ptr_ = app_disp_ptr;
      init_ = true;
    }
  }

  return init_;
}

void excel_loader::uninit() {
  if (init_) {
    DISPPARAMS params = { 0 };
    VARIANT result = { 0 };
    VariantInit(&result);

    if (wbs_disp_ptr_ != nullptr) {
      wbs_disp_ptr_->Invoke(id_[DISPID_WBS_CLOSE],
        IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params,
        &result, nullptr, nullptr);
      wbs_disp_ptr_ = nullptr;
    }

    if (app_disp_ptr_ != nullptr) {
      app_disp_ptr_->Invoke(id_[DISPID_APP_QUIT],
        IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params,
        &result, nullptr, nullptr);
      app_disp_ptr_ = nullptr;
    }
    init_ = false;
  }
}

bool excel_loader::set_file(const wchar_t *file) {
  return open_params_.set_filename(file);
}

bool excel_loader::try_open(const wchar_t *passwd) {
  bool success = false;

  if (init_ && passwd != nullptr) {
    VARIANT result = { 0 };
    VariantInit(&result);
    open_params_.set_password(passwd);
    HRESULT hr = wbs_disp_ptr_->Invoke(id_[DISPID_WBS_OPEN],
      IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, open_params_,
      &result, nullptr, nullptr);
    if (SUCCEEDED(hr)) {
      result.pdispVal->Release();
      success = true;
    }
  }

  return success;
}

HRESULT excel_loader::create_excel_instance(IDispatch** disp_pp) const {
  CLSID clsid = { 0 };
  HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
  if (SUCCEEDED(hr) && disp_pp != nullptr) {
    hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER,
      IID_IDispatch, reinterpret_cast<void**>(disp_pp));
    if (SUCCEEDED(hr)) {
      hr = OleRun(*disp_pp);
    }
  }

  return hr;
}

DISPID excel_loader::get_dispid(IDispatch* disp_ptr, const wchar_t *name) const {
  DISPID id = -1;
  if (disp_ptr != nullptr) {
    disp_ptr->GetIDsOfNames(IID_NULL, const_cast<wchar_t**>(&name),
      1, LOCALE_USER_DEFAULT, &id);
  }

  return id;
}
