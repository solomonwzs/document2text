#include "utils/utils.h"

#include <dirent.h>
#include <errno.h>
#include <fcntl.h>
#include <pthread.h>
#include <stdio.h>
#include <string.h>
#include <sys/stat.h>
#include <sys/time.h>
#include <time.h>
#include <unistd.h>

#include <algorithm>
#include <functional>
#include <memory>
#include <string>
#include <vector>

#define CONCAT_(_A, _B) _A##_B
#define CONCAT(_A, _B)  CONCAT_(_A, _B)
#define _defer(_fn_) \
  std::shared_ptr<void> CONCAT(__defer, __LINE__)(nullptr, _fn_)

#define _STYLE_Debug "\e[3;36m"
#define _STYLE_Info  "\e[3;32m"
#define _STYLE_Warn  "\e[3;33m"
#define _STYLE_Err   "\e[3;31m"
#define xmlog(_type_, _fmt_, ...)                                     \
  printf(_STYLE_##_type_ "%.1s [%s:%s:%d]\e[0m " _fmt_ "\n", #_type_, \
         __FILE__, __func__, __LINE__, ##__VA_ARGS__)

namespace utils {

template <typename Bytes>
static int __read_file(const char* filename, Bytes* buff) {
  int fd = open(filename, O_RDONLY);
  if (fd == -1) {
    xmlog(Err, "open: %s, %s", filename, strerror(errno));
    return -1;
  }
  _defer([&](...) { close(fd); });

  struct stat st;
  if (fstat(fd, &st) != 0) {
    xmlog(Err, "fstat: %s, %s", filename, strerror(errno));
    return -1;
  }

  buff->resize(st.st_size);
  ssize_t n = read(fd, &(*buff)[0], buff->size());
  if (n < 0) {
    xmlog(Err, "read: %s, %s", filename, strerror(errno));
    return -1;
  } else if (n != st.st_size) {
    xmlog(Err, "read size error, except %ld, get %ld", st.st_size, n);
    return -1;
  }
  return 0;
}

int read_file(const char* filename, std::vector<char>* data) {
  return __read_file(filename, data);
}

int read_file(const char* filename, std::string* data) {
  return __read_file(filename, data);
}

// =============================================================================

int UTF8String::CountUTF8WordWidth(const char* c, size_t wlen) {
  if (*c >= 0 && *c <= 127) {
    return 1;
  } else if ((*c & 0xE0) == 0xC0) {
    return 2;
  } else if ((*c & 0xF0) == 0xE0) {
    return 3;
  } else if ((*c & 0xF8) == 0xF0) {
    return 4;
  } else {
    return -1;
  }
}

size_t UTF8String::CountUTF8WordCnt(const char* s, size_t slen) {
  const char* p = s;
  size_t len = 0;
  for (size_t i = 0; i < slen;) {
    int width = CountUTF8WordWidth(p + i, slen - i);
    if (width < 0) {
      i += 1;
      len += 1;
    } else {
      i += width;
      len += 1;
    }
  }
  return len;
}

size_t UTF8String::CountUTF8WordCnt(const std::string& str) {
  return CountUTF8WordCnt(str.c_str(), str.length());
}

size_t UTF8String::FixUTF8WordCnt(const char* s, size_t cnt) {
  size_t offset = 0;
  for (size_t i = 0; i < cnt; ++i) {
    int w = CountUTF8WordWidth(s + offset);
    offset += w == -1 ? 1 : w;
  }
  return offset;
}

}  // namespace utils
