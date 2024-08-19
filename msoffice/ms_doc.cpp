#include "msoffice/ms_doc.h"

#include <string.h>

#include "utils/utils.h"

#define CONCAT_(_A, _B) _A##_B
#define CONCAT(_A, _B) CONCAT_(_A, _B)
#define _defer(_fn_) \
  std::shared_ptr<void> CONCAT(__defer, __LINE__)(nullptr, _fn_)

namespace msoffice {

namespace doc {

static const std::string g_0TableDirName = "0Table";
static const std::string g_1TableDirName = "1Table";
static const std::string g_WordDocDirName = "WordDocument";

int PlcPcd::ParseFrom(const FibRgFcLcb97_t &fib_rg_fc_lcb97,
                      const std::vector<char> &table_stream) {
  const char *clx_ptr = table_stream.data() + fib_rg_fc_lcb97.fcClx;
  const char *pcdt_ptr = nullptr;

  const char *clx_m_ptr = clx_ptr;
#define clxt reinterpret_cast<const uint8_t *>(clx_m_ptr)
  for (; clx_m_ptr < table_stream.data() + table_stream.size() &&
         *clxt == 0x01;) {
    auto cbGrpprl = reinterpret_cast<const int16_t *>(clx_m_ptr + 1);
    clx_m_ptr += 1 + 2 + *cbGrpprl;
  }

  if (clx_m_ptr < table_stream.data() + table_stream.size()) {
    if (*clxt == 0x02) {
      pcdt_ptr = clx_ptr;
    } else if (*clxt == 0x01) {
      auto cb_grpprl = reinterpret_cast<const int16_t *>(clxt + 1);
      pcdt_ptr = clx_ptr + 1 + *cb_grpprl;
    } else {
      return -1;
    }
  } else {
    return -1;
  }
#undef clxt

  if (pcdt_ptr - table_stream.data() + 1 + sizeof(uint32_t) >
      table_stream.size()) {
    return -1;
  }
  auto lcb = reinterpret_cast<const uint32_t *>(pcdt_ptr + 1);
  if (clx_ptr + fib_rg_fc_lcb97.lcbClx !=
          pcdt_ptr + 1 + sizeof(uint32_t) + *lcb ||
      (*lcb - sizeof(int32_t)) % (sizeof(int32_t) + sizeof(Pcd_t)) != 0) {
    return -1;
  }

  size_t cp_cnt =
      (*lcb - sizeof(int32_t)) / (sizeof(int32_t) + sizeof(Pcd_t)) + 1;
  size_t pcd_cnt = cp_cnt - 1;

  m_cp.resize(cp_cnt);
  memcpy(m_cp.data(), pcdt_ptr + 1 + sizeof(uint32_t),
         sizeof(int32_t) * cp_cnt);

  m_pcd.resize(pcd_cnt);
  memcpy(m_pcd.data(),
         pcdt_ptr + 1 + sizeof(uint32_t) + sizeof(int32_t) * cp_cnt,
         sizeof(Pcd_t) * pcd_cnt);

  return 0;
}

// =============================================================================

const uint16_t MsDOC::_fib_base_wIdent = 0xA5EC;
const size_t MsDOC::_fib_rg_fc_lcb97_offset = 0x9A;

int MsDOC::ParseFromFile(const std::string &filename) {
  std::vector<char> data;
  utils::read_file(filename.c_str(), &data);
  if (data.size() < sizeof(compound_doc_header_t)) {
    return -1;
  }
  return ParseFromBytes(std::move<>(data));
}

int MsDOC::parse() {
  m_idx_tab0 = -1;
  m_idx_tab1 = -1;
  m_idx_word_doc = -1;

  int len_tab0 = -1;
  int len_tab1 = -1;
  int len_word_doc = -1;

  auto &dirs = m_comp_doc.GetDirEntries();
  for (int i = 0; i < static_cast<int>(dirs.size()); ++i) {
    if (dirs[i].type == kDirEntryTypeEmpty) {
      continue;
    }
    std::string dirname;
    if (Utf16ToUtf8(dirs[i].unicode_name,
                    dirs[i].unicode_name + dirs[i].name_len / 2 - 1,
                    &dirname) != 0) {
      continue;
    }
    if (dirname == g_0TableDirName && len_tab0 < dirs[i].size_of_x) {
      m_idx_tab0 = i;
      len_tab0 = dirs[i].size_of_x;
    } else if (dirname == g_1TableDirName && len_tab1 < dirs[i].size_of_x) {
      m_idx_tab1 = i;
      len_tab1 = dirs[i].size_of_x;
    } else if (dirname == g_WordDocDirName &&
               len_word_doc < dirs[i].size_of_x) {
      m_idx_word_doc = i;
      len_word_doc = dirs[i].size_of_x;
    }
  }
  if (m_idx_word_doc == -1 || (m_idx_tab0 == -1 && m_idx_tab1 == -1)) {
    return -1;
  }

  return 0;
}

int MsDOC::FetchText(const fetch_text_options_t *opts,
                     std::string *text) const {
  auto &dirs = m_comp_doc.GetDirEntries();

  std::vector<char> word_doc_stream;
  if (m_comp_doc.GetDirEntryStream(dirs[m_idx_word_doc], &word_doc_stream) !=
          0 ||
      word_doc_stream.size() < sizeof(fib_base_t)) {
    return -1;
  }

  auto fib_base = reinterpret_cast<fib_base_t *>(word_doc_stream.data());
  if (fib_base->wIdent != _fib_base_wIdent || fib_base->fEncrypted()) {
    return -1;
  }

  if (word_doc_stream.size() < _fib_rg_fc_lcb97_offset ||
      word_doc_stream.size() - _fib_rg_fc_lcb97_offset <
          sizeof(FibRgFcLcb97_t)) {
    return -1;
  }

  auto fib_rg_fc_lcb97 = reinterpret_cast<FibRgFcLcb97_t *>(
      word_doc_stream.data() + _fib_rg_fc_lcb97_offset);

  std::vector<char> table_stream;
  if (m_comp_doc.GetDirEntryStream(
          dirs[fib_base->fWhichTblStm() == 0 ? m_idx_tab0 : m_idx_tab1],
          &table_stream) != 0) {
    return -1;
  }

  PlcPcd plc_pcd;
  if (plc_pcd.ParseFrom(*fib_rg_fc_lcb97, table_stream) != 0) {
    return -1;
  }

  size_t max_fetch_text_len =
      opts != nullptr ? opts->max_fetch_text_len
                      : __defaultFetchTextOptions.max_fetch_text_len;
  auto &cp_list = plc_pcd.GetCP();
  auto &pcd_list = plc_pcd.GetPcd();
  for (size_t i = 0; i < pcd_list.size() && max_fetch_text_len > 0; ++i) {
    size_t offset;
    size_t len = cp_list[i + 1] - cp_list[i];
    if (len > max_fetch_text_len) {
      len = max_fetch_text_len;
    }
    if (pcd_list[i].fc.fCompressed() == 1) {  // ANSI
      offset = pcd_list[i].fc.fc() / 2;
      if (word_doc_stream.size() < offset ||
          word_doc_stream.size() - offset < len) {
        return -1;
      }
      text->append(word_doc_stream.data() + offset, len);
    } else {  // Unicode
      offset = pcd_list[i].fc.fc();
      if (word_doc_stream.size() < offset ||
          word_doc_stream.size() - offset < len * 2) {
        return -1;
      }
      auto ptr = reinterpret_cast<char16_t *>(word_doc_stream.data() + offset);

      std::string s;
      if (Utf16ToUtf8(ptr, ptr + len, &s) != 0) {
        return -1;
      } else {
        text->append(std::move(s));
      }
    }

    max_fetch_text_len -= len;
  }

  RemoveControlCharacter(text);

  return 0;
}

}  // namespace doc

}  // namespace msoffice
