#include "simplepdf/simplepdf.h"

#include <poppler/GlobalParams.h>
#include <poppler/TextOutputDev.h>
#include <stdio.h>

#include <chrono>
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

  bool radialShadedFill(GfxState* /*state*/, GfxRadialShading* /*shading*/,
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

void Init(const std::string& poppler_data_dir) {
  globalParams =
      std::unique_ptr<GlobalParams>(new GlobalParams(poppler_data_dir));
}

SimplePDF::SimplePDF(const char* buf, size_t buf_len) : m_doc(nullptr) {
  auto mem = std::make_unique<MemStream>(buf, 0, buf_len, Object::null());
  if (mem == nullptr) {
    return;
  }
  m_doc = new PDFDoc(std::move(mem));
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
    Page* page = m_doc->getPage(i);
    ImageOutputDev out;
    page->displaySlice(&out, 72, 72, 0, false, false, -1, -1, -1, -1, false);
  }
}

static bool annot_display_decide_cbk(Annot* annot, void*) {
  return false;
}

struct abort_chk_t {
  std::chrono::time_point<std::chrono::high_resolution_clock> start;
  bool abort;
  uint32_t cnt;
};
static bool abort_check_cbk(void* ud) {
  auto abort_chk = reinterpret_cast<abort_chk_t*>(ud);
  abort_chk->cnt += 1;
  // abort_chk->abort = time(nullptr) - abort_chk->start > 2;
  abort_chk->abort =
      abort_chk->cnt > 2000 ||
      std::chrono::duration_cast<std::chrono::milliseconds>(
          std::chrono::high_resolution_clock::now() - abort_chk->start)
              .count() > 500;
  return abort_chk->abort;
}

GooString SimplePDF::PageText(int n) {
  if (n < 1 || n > PagesCnt()) {
    return GooString("");
  }

  Page* page = m_doc->getPage(n);
  if (!page->isOk()) {
    return GooString("");
  }

  abort_chk_t abort_chk;
  abort_chk.start = std::chrono::high_resolution_clock::now();
  abort_chk.abort = false;
  abort_chk.cnt = 0;

  _SimpleTextOutputDev out;
  page->displaySlice(&out, 72, 72, 0, false, false, -1, -1, -1, -1, false,
                     abort_check_cbk, &abort_chk, annot_display_decide_cbk,
                     nullptr);

  double w = page->getMediaWidth();
  double h = page->getMediaHeight();
  return out.getText(PDFRectangle(0,0, w, h));
}

}  // namespace simplepdf
