#include "msoffice/utils.h"

#include <stdio.h>

#include <codecvt>
#include <locale>

#define _STYLE_Impt "\e[3;35m"
#define _STYLE_Info "\e[3;32m"
#define _STYLE_Err "\e[3;31m"
#define _STYLE_Warn "\e[3;33m"
#define _STYLE_Debug "\e[3;36m"
#define xmlog(_type_, _fmt_, ...)                                     \
  printf(_STYLE_##_type_ "%.1s [%s:%s:%d]\e[0m " _fmt_ "\n", #_type_, \
         __FILE__, __func__, __LINE__, ##__VA_ARGS__)

namespace msoffice {

const fetch_text_options_t __defaultFetchTextOptions;

void RemoveControlCharacter(std::string *text) {
  for (auto &ch : *text) {  // Remove control character
    if (ch == '\r' || ch == '\n') {
      ch = '\n';
    } else if ((1 <= ch && ch <= 31) || ch == 127) {
      ch = ' ';
    }
  }
}

int Utf16ToUtf8(const char16_t *begin, const char16_t *end, std::string *u8) {
  std::wstring_convert<std::codecvt_utf8_utf16<char16_t>, char16_t> convert;
  try {
    *u8 = convert.to_bytes(begin, end);
    return 0;
  } catch (const std::exception &e) {
    xmlog(Err, "u16_to_u8 err, %s", e.what());
    return -1;
  }
}

}  // namespace msoffice
