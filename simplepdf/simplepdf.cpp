#include "simplepdf/simplepdf.h"

#include <poppler/GlobalParams.h>
#include <poppler/TextOutputDev.h>
#include <stdio.h>

#include <memory>

#include "simplepdf/imgoutputdev.h"

#define _STYLE_Debug "\e[3;36m"
#define _STYLE_Info "\e[3;32m"
#define _STYLE_Warn "\e[3;33m"
#define _STYLE_Err "\e[3;31m"
#define slog(_type_, _fmt_, ...)                                      \
  printf(_STYLE_##_type_ "%.1s [%s:%s:%d]\e[0m " _fmt_ "\n", #_type_, \
         __FILE__, __func__, __LINE__, ##__VA_ARGS__)

namespace simplepdf {

class _SimpleTextOutputDev : public TextOutputDev {
 public:
  _SimpleTextOutputDev() : TextOutputDev(nullptr, false, 0, false, false) {}
  virtual ~_SimpleTextOutputDev() {}

  bool radialShadedFill(GfxState * /*state*/, GfxRadialShading * /*shading*/,
                        double /*sMin*/, double /*sMax*/) override {
    return true;
  }

  bool useShadedFills(int type) override {
    return type == 3;
  }
  // void drawImage(GfxState *state, Object *ref, Stream *str, int width,
  //                int height, GfxImageColorMap *colorMap, bool interpolate,
  //                const int *maskColors, bool inlineImg) override {}
};

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

  _SimpleTextOutputDev out;
  page->displaySlice(&out, 72, 72, 0, false, false, -1, -1, -1, -1, false);
  double w = page->getMediaWidth();
  double h = page->getMediaHeight();
  GooString *text = out.getText(0, 0, w, h);
  return std::unique_ptr<GooString>(text);
}

}  // namespace simplepdf
