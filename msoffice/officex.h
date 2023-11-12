#pragma once

#include <stdint.h>
#include <zip.h>

#include <map>
#include <string>
#include <vector>

namespace msoffice {

namespace officex {

class ZipHelper {
 public:
  ZipHelper();
  ~ZipHelper();

  ZipHelper(const ZipHelper &) = delete;
  ZipHelper &operator=(const ZipHelper &) = delete;

  int OpenFromBytes(const std::vector<char> &data);
  int OpenFromBytes(const char *data, size_t data_len);

  int ReadByName(const std::string &name, std::vector<char> *data);
  int ReadByName(const std::string &name, std::string *data);

  inline const std::map<std::string, int64_t> &GetName2Idx() const {
    return m_name2idx;
  }

 private:
  zip_source_t *m_zsrc;
  zip_t *m_zfd;
  std::map<std::string, int64_t> m_name2idx;
};

int MsDOCxFetchText(ZipHelper &zip, size_t max_len, std::string *text);
int MsDOCxFetchText(const char *xml_text, size_t xml_len, size_t max_len,
                    std::string *text);

int MsPPTxFetchText(ZipHelper &zip, size_t max_len, std::string *text);
int MsPPTxFetchText(const char *xml_text, size_t xml_len, size_t max_len,
                    std::string *text, size_t *fetch_len = nullptr);

struct xlsx_sheet_bar_t {
  std::string id;
  std::string name;
  std::string state;
};

int MsXLSxFetchSheetBarList(ZipHelper &zip,
                            std::vector<xlsx_sheet_bar_t> *sheets);
int MsXLSxFetchSST(ZipHelper &zip, int max_sst_cnt,
                   std::vector<std::string> *sst);
int MsXLSxFetchText(ZipHelper &zip, size_t max_len, int max_sst_cnt,
                    const std::string &delimiter, std::string *text);

}  // namespace officex

}  // namespace msoffice
