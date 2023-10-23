#pragma once

#include <stdio.h>
#include <time.h>

#include <chrono>
#include <string>
#include <vector>

#define _STYLE_Info  "\e[3;32m"
#define _STYLE_Err   "\e[3;31m"
#define _STYLE_Warn  "\e[3;33m"
#define _STYLE_Debug "\e[3;36m"
#define xmlog(_type_, _fmt_, ...)                                     \
  printf(_STYLE_##_type_ "%.1s [%s:%s:%d]\e[0m " _fmt_ "\n", #_type_, \
         __FILE__, __func__, __LINE__, ##__VA_ARGS__)

namespace utils {

int read_file(const char* filename, std::vector<char>* data);
int read_file(const char* filename, std::string* data);

class UTF8String {
 public:
  static size_t CountUTF8WordCnt(const std::string& str);
  static size_t CountUTF8WordCnt(const char* s, size_t slen);
  static int CountUTF8WordWidth(const char* word, size_t wlen = 4);

  static size_t FixUTF8WordCnt(const char* s, size_t cnt);
};

}  // namespace utils

#undef xmlog
