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

#include "bruteforce_excel/worker.h"
#include <cstdint>
#include <iostream>
#include <vector>
#include <mutex>  // NOLINT
#include <thread>  // NOLINT
#include "bruteforce_excel/excel_loader.h"
#include "bruteforce_excel/tick_counter.h"

bool worker::chk_excel_installed() const {
  excel_loader excel;
  return excel.init();
}

bool worker::run(const wchar_t *filename) const {
  enum { PASSWORD_BUFLEN = 128 };

  work_info_t wi = { filename };
  const int32_t num_of_thread = std::thread::hardware_concurrency();

  std::vector<std::thread*> thread_pool(num_of_thread);

  tick_counter tck_cnt;
  for (auto &thread_ptr : thread_pool) {
    thread_ptr = new std::thread(std::bind(&thread_proc, std::ref(wi)));
  }
  for (auto &thread_ptr : thread_pool) {
    if (thread_ptr != nullptr) {
      thread_ptr->join();
      delete thread_ptr;
    }
  }

  const double aps = wi.cnt_ / (tck_cnt.get_elapsed() / 1000.0);
  std::cout << "password attempts per second: " << aps << std::endl;

  return wi.find_;
}

void worker::thread_proc(work_info_t &wi) {
  enum { PASSWORD_BUFLEN = 128 };

  excel_loader excel;
  if (excel.init() && excel.set_file(wi.filename_)) {
    bool get_pwd = false;
    wchar_t password[PASSWORD_BUFLEN] = { 0 };
    do {
      {
        std::lock_guard<std::mutex> lock_holer(wi.mtx_);
        get_pwd = std::wcin.getline(password, _countof(password)).good();
      }

      if (get_pwd && excel.try_open(password)) {
        std::wcout << L"password: " << password << std::endl;
        wi.find_ = true;
      }
      ++wi.cnt_;
    } while (wi.find_ == false && get_pwd == true);
  }
}
