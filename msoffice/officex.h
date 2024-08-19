#pragma once

#include <stdint.h>
#include <zip.h>

#include <map>
#include <string>
#include <vector>

#include "msoffice/utils.h"

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

  int ReadByName(const std::string &name, size_t max_read_len,
                 std::vector<char> *data);
  int ReadByName(const std::string &name, size_t max_read_len,
                 std::string *data);

  inline const std::map<std::string, int64_t> &GetName2Idx() const {
    return m_name2idx;
  }

 private:
  zip_source_t *m_zsrc;
  zip_t *m_zfd;
  std::map<std::string, int64_t> m_name2idx;
};

int MsDOCxFetchText(ZipHelper &zip, const fetch_text_options_t *opts,
                    std::string *text);
int MsDOCxFetchText(char *xml_text, const fetch_text_options_t *opts,
                    std::string *text);

int MsPPTxFetchText(ZipHelper &zip, const fetch_text_options_t *opts,
                    std::string *text);
int MsPPTxFetchText(char *xml_text, const fetch_text_options_t *opts,
                    std::string *text, size_t *fetch_len = nullptr);

struct xlsx_sheet_bar_t {
  std::string rid;
  std::string name;
  std::string state;
};

int MsXLSxFetchRelationships(ZipHelper &zip, size_t xml_max_file_len,
                             std::map<std::string, std::string> *rid2target);

int MsXLSxFetchSheetBarList(ZipHelper &zip, size_t xml_max_file_len,
                            std::vector<xlsx_sheet_bar_t> *sheets,
                            bool *is_xtag = nullptr);

int MsXLSxFetchSST(ZipHelper &zip, size_t xml_max_file_len, int max_sst_cnt,
                   std::vector<std::string> *sst);
int MsXLSxFetchText(ZipHelper &zip, const fetch_text_options_t *opts,
                    std::string *text);

}  // namespace officex

}  // namespace msoffice
