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

#ifndef BRUTEFORCE_EXCEL_TICK_COUNTER_H_
#define BRUTEFORCE_EXCEL_TICK_COUNTER_H_

#include <cstdint>
#include <chrono>  // NOLINT

class tick_counter final {
 public:
  typedef std::chrono::high_resolution_clock clock;
  typedef std::chrono::milliseconds msec;

  tick_counter() { reset(); }
  ~tick_counter() = default;

  inline void reset() { base_ = clock::now(); }
  inline uint64_t get_elapsed() const {
    return std::chrono::duration_cast<msec>(clock::now() - base_).count(); }

 private:
  std::chrono::high_resolution_clock::time_point base_;
};

#endif  // BRUTEFORCE_EXCEL_TICK_COUNTER_H_
