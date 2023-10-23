#include "simplepdf/simplepdf.h"

#include <poppler/GlobalParams.h>
#include <poppler/TextOutputDev.h>
#include <stdio.h>

#include <map>
#include <memory>
#include <string>
#include <vector>

#include "simplepdf/imgoutputdev.h"

#define _STYLE_Debug "\e[3;36m"
#define _STYLE_Info  "\e[3;32m"
#define _STYLE_Warn  "\e[3;33m"
#define _STYLE_Err   "\e[3;31m"
#define xmlog(_type_, _fmt_, ...)                                     \
  printf(_STYLE_##_type_ "%.1s [%s:%s:%d]\e[0m " _fmt_ "\n", #_type_, \
         __FILE__, __func__, __LINE__, ##__VA_ARGS__)

namespace simplepdf {

void Init(const char *poppler_data_dir) {
  globalParams =
      std::unique_ptr<GlobalParams>(new GlobalParams(poppler_data_dir));
}

SimplePDF::SimplePDF(const char *buf, size_t buf_len) : m_doc(nullptr) {
  MemStream *mem = new MemStream(buf, 0, buf_len, Object(objNull));
  if (mem == nullptr) {
    return;
  }
  m_doc = new PDFDoc(mem);
}

SimplePDF::~SimplePDF() {
  if (m_doc != nullptr) {
    delete m_doc;
  }
}

bool SimplePDF::IsOK() const {
  return m_doc != nullptr && m_doc->isOk();
}

int SimplePDF::PagesCnt() const {
  return IsOK() ? m_doc->getNumPages() : 0;
}

void SimplePDF::Debug() {
  for (int i = 1; i <= m_doc->getNumPages(); ++i) {
    Page *page = m_doc->getPage(i);
    ImageOutputDev out;
    page->displaySlice(&out, 72, 72, 0, false, false, -1, -1, -1, -1, false);
  }
}

std::unique_ptr<GooString> SimplePDF::PageText(int n) {
  if (n < 1 || n > PagesCnt()) {
    return nullptr;
  }

  Page *page = m_doc->getPage(n);
  if (!page->isOk()) {
    return nullptr;
  }

  TextOutputDev out(nullptr, false, 0, false, false);
  page->displaySlice(&out, 72, 72, 0, false, false, -1, -1, -1, -1, false);
  double w = page->getMediaWidth();
  double h = page->getMediaHeight();
  GooString *text = out.getText(0, 0, w, h);
  return std::unique_ptr<GooString>(text);
}

}  // namespace simplepdf
