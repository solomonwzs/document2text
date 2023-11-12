#pragma once

#include <stdio.h>
#include <time.h>

#include <chrono>
#include <string>
#include <vector>

namespace utils {

int read_file(const char* filename, std::vector<char>* data);
int read_file(const char* filename, std::string* data);

size_t count_utf8_word_cnt(const std::string& str);
size_t count_utf8_word_cnt(const char* s, size_t slen);

size_t fix_utf8_word_cnt(const char* s, size_t cnt);

}  // namespace utils
