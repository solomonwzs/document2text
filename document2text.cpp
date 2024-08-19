#include <stdio.h>
#include <stdlib.h>

#include <string>
#include <vector>

#include "msoffice/ms_doc.h"
#include "msoffice/ms_ppt.h"
#include "msoffice/ms_xls/ms_xls.h"
#include "msoffice/officex.h"
#include "simplepdf/simplepdf.h"
#include "utils/utils.h"

enum document_type_t {
  kDocTypeUnknown = 0,
  kDocTypePDF = 1,
  kDocTypeDOC = 2,
  kDocTypePPT = 3,
  kDocTypeXLS = 4,
  kDocTypeDOCX = 5,
  kDocTypePPTX = 6,
  kDocTypeXLSX = 7,
};

enum doc2txt_result_t {
  kDoc2txtOK = 0,
  kDoc2txtFail = -1,
  kDoc2txtConvertErr = -2,
};

struct fetch_opts_t {
  size_t max_fetch_text_len;
  int max_fetch_pdf_page_cnt;
  int max_xls_sst_cnt;
  document_type_t type;
};

static bool is_document_pdf(const char *p, size_t plen) {
  if (plen > 5 && p != nullptr && strncmp(p, "%PDF-", 5) == 0) {
    return true;
  }
  std::vector<char> buf(128);
  size_t len = std::min(plen, buf.size() - 1);
  memcpy(buf.data(), p, len);
  buf[len] = 0;
  for (size_t i = 0; i < len; ++i) {
    if (buf[i] == 0) {
      buf[i] = 1;
    }
  }
  return strstr(buf.data(), "%PDF-") != nullptr;
}

static inline bool is_zip(const char *p, size_t plen) {
  return plen > 30 && *reinterpret_cast<const uint32_t *>(p) == 0x04034b50;
}

static doc2txt_result_t pdf2text(const char *data, size_t len,
                                 size_t max_fetch_text_len,
                                 int max_fetch_pdf_page_cnt,
                                 std::string *text) {
  simplepdf::SimplePDF pdf(data, len);
  int page_cnt = pdf.PagesCnt();
  for (int i = 1;
       i <= page_cnt && i <= max_fetch_pdf_page_cnt && max_fetch_text_len > 0;
       ++i) {
    try {
      auto t = pdf.PageText(i);
      if (t == nullptr || t->c_str() == nullptr) {
        continue;
      }

      size_t len = utils::count_utf8_word_cnt(t->c_str(), t->getLength());
      if (max_fetch_text_len >= len) {
        text->append(t->c_str(), t->getLength());
        max_fetch_text_len -= len;
      } else {
        size_t offset =
            utils::fix_utf8_word_cnt(t->c_str(), max_fetch_text_len);
        text->append(t->c_str(), offset);
        max_fetch_text_len = 0;
      }
    } catch (std::exception &ex) {
    }
  }
  return kDoc2txtOK;
}

doc2txt_result_t document2text(const char *data, size_t len,
                               const fetch_opts_t &opts, std::string *text,
                               document_type_t *type) {
  *type = kDocTypeUnknown;

  if (opts.type == kDocTypePDF ||
      (opts.type == kDocTypeUnknown && is_document_pdf(data, len))) {
    *type = kDocTypePDF;
    return pdf2text(data, len, opts.max_fetch_text_len,
                    opts.max_fetch_pdf_page_cnt, text);
  }

  msoffice::fetch_text_options_t fopts;
  fopts.max_fetch_text_len = opts.max_fetch_text_len;
  fopts.fetch_text_from_drawing = true;
  fopts.xls_delimiter = ",";
  fopts.xls_skip_blank_cell = true;
  fopts.xls_max_sst_cnt = opts.max_xls_sst_cnt;
  fopts.xml_max_file_len = 1024 * 1024;

  if (is_zip(data, len)) {
    msoffice::officex::ZipHelper zip;
    if (zip.OpenFromBytes(data, len) != 0) {
      return kDoc2txtFail;
    }

    *type = opts.type;
    if (*type == kDocTypeUnknown) {
      auto n2i = zip.GetName2Idx();
      if (n2i.find("word/document.xml") != n2i.end()) {
        *type = kDocTypeDOCX;
      } else if (n2i.find("ppt/presentation.xml") != n2i.end()) {
        *type = kDocTypePPTX;
      } else if (n2i.find("xl/workbook.xml") != n2i.end()) {
        *type = kDocTypeXLSX;
      } else {
        return kDoc2txtFail;
      }
    }

    if (*type == kDocTypeDOCX) {
      return msoffice::officex::MsDOCxFetchText(zip, &fopts, text) == 0
                 ? kDoc2txtOK
                 : kDoc2txtConvertErr;
    } else if (*type == kDocTypePPTX) {
      return msoffice::officex::MsPPTxFetchText(zip, &fopts, text) == 0
                 ? kDoc2txtOK
                 : kDoc2txtConvertErr;
    } else if (*type == kDocTypeXLSX) {
      return msoffice::officex::MsXLSxFetchText(zip, &fopts, text) == 0
                 ? kDoc2txtOK
                 : kDoc2txtConvertErr;
    } else {
      return kDoc2txtFail;
    }
  }

  msoffice::CompoundDocument comp_doc;
  if (comp_doc.ParseFromBytes(data, len) != 0) {
    return kDoc2txtFail;
  }

  *type = opts.type;
  if (*type == kDocTypeUnknown) {
    for (auto &i : comp_doc.GetDirEntries()) {
      std::string dirname;
      if (i.type == msoffice::kDirEntryTypeEmpty ||
          msoffice::Utf16ToUtf8(i.unicode_name,
                                i.unicode_name + i.name_len / 2 - 1,
                                &dirname) != 0) {
        continue;
      }
      if (dirname == "WordDocument") {
        *type = kDocTypeDOC;
        break;
      } else if (dirname == "PowerPoint Document") {
        *type = kDocTypePPT;
        break;
      } else if (dirname == "Workbook") {
        *type = kDocTypeXLS;
        break;
      }
    }
  }

  if (*type == kDocTypeDOC) {
    msoffice::doc::MsDOC doc;
    if (doc.ParseFromCompoundDocument(std::move(comp_doc)) != 0 ||
        doc.FetchText(&fopts, text) != 0) {
      return kDoc2txtConvertErr;
    }
  } else if (*type == kDocTypePPT) {
    msoffice::ppt::MsPPT ppt;
    if (ppt.ParseFromCompoundDocument(std::move(comp_doc)) != 0 ||
        ppt.FetchText(&fopts, text) != 0) {
      return kDoc2txtConvertErr;
    }
  } else if (*type == kDocTypeXLS) {
    msoffice::xls::MsXLS xls;
    if (xls.ParseFromCompoundDocument(std::move(comp_doc)) != 0 ||
        xls.FetchText(&fopts, text) != 0) {
      return kDoc2txtConvertErr;
    }
  } else {
    return kDoc2txtFail;
  }
  return kDoc2txtOK;
}

int main(int argc, char **argv) {
  assert(argc >= 2);
  const char *filename = argv[1];

  simplepdf::Init(nullptr);

  std::vector<char> data;
  assert(utils::read_file(filename, &data) == 0);
  fetch_opts_t opts = {
      .max_fetch_text_len = 40960,
      .max_fetch_pdf_page_cnt = 20,
  };
  std::string text;
  document_type_t type;
  assert(document2text(data.data(), data.size(), opts, &text, &type) == 0);
  printf("%s", text.c_str());
  return 0;
}
