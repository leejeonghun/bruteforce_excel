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

#ifndef BRUTEFORCE_EXCEL_WORKER_H_
#define BRUTEFORCE_EXCEL_WORKER_H_

#include <atomic>
#include <mutex>  // NOLINT

class worker final {
 public:
  worker() = default;
  ~worker() = default;

  bool chk_excel_installed() const;
  bool run(const wchar_t *filename) const;

 private:
  struct work_info_t {
    const wchar_t *filename_;
    std::mutex mtx_;
    std::atomic_bool find_;
    std::atomic_uint64_t cnt_;
  };

  static void thread_proc(work_info_t &wi);  // NOLINT
};

#endif  // BRUTEFORCE_EXCEL_WORKER_H_
