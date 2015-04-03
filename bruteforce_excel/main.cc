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

#include <shlwapi.h>
#include <iostream>
#include "bruteforce_excel/worker.h"

#pragma comment(lib, "shlwapi")

int wmain(int argc, wchar_t* argv[]) {
  const char usage[] = "Usage : bruteforce_excel.exe "
    "sample.xlsx < dictionary.txt";
  const char init_err[] = "Initialization failed, "
    "check Microsoft Excel is installed.";

  int retcode = -1;

  if (argc == 1 || argc > 1 && PathFileExists(argv[1]) == FALSE) {
    std::cout << usage << std::endl;
  } else {
    worker brute_force;
    if (brute_force.chk_excel_installed()) {
      retcode = brute_force.run(argv[1]) ? 0 : -2;
    } else {
      std::cerr << init_err << std::endl;
    }
  }

  return retcode;
}
