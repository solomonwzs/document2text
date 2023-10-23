#pragma once

#include <poppler/GlobalParams.h>
#include <poppler/PDFDoc.h>

#include <memory>

namespace simplepdf {

void Init(const char *poppler_data_dir = nullptr);

class SimplePDF {
 public:
  SimplePDF(const char *buf, size_t buf_len);
  ~SimplePDF();

  bool IsOK() const;
  int PagesCnt() const;
  std::unique_ptr<GooString> PageText(int n);

  void Debug();

 private:
  PDFDoc *m_doc;
};

}  // namespace simplepdf
