#pragma once

#include <unistd.h>

#include <limits>
#include <string>

namespace msoffice {

struct fetch_text_options_t {
  size_t max_fetch_text_len = std::numeric_limits<size_t>::max();
  bool fetch_text_from_drawing = false;
  std::string xls_delimiter = ",";
  bool xls_skip_blank_cell = true;
  int xls_max_sst_cnt = 0xffff;
  size_t xml_max_file_len = 1024 * 1024;
};

extern const fetch_text_options_t __defaultFetchTextOptions;

void RemoveControlCharacter(std::string* text);

int Utf16ToUtf8(const char16_t* begin, const char16_t* end, std::string* u8);

}  // namespace msoffice
