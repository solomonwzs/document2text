#include "msoffice/compound_document.h"

#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include <codecvt>
#include <locale>
#include <memory>
#include <numeric>

#include "utils/utils.h"

#define _STYLE_Info  "\e[3;32m"
#define _STYLE_Err   "\e[3;31m"
#define _STYLE_Warn  "\e[3;33m"
#define _STYLE_Debug "\e[3;36m"
#define slog(_type_, _fmt_, ...)                                      \
  printf(_STYLE_##_type_ "%.1s [%s:%s:%d]\e[0m " _fmt_ "\n", #_type_, \
         __FILE__, __func__, __LINE__, ##__VA_ARGS__)

#define CONCAT_(_A, _B) _A##_B
#define CONCAT(_A, _B)  CONCAT_(_A, _B)
#define _defer(_fn_) \
  std::shared_ptr<void> CONCAT(__defer, __LINE__)(nullptr, _fn_)

namespace msoffice {

std::string vector_str(const std::vector<int> &vec) {
  std::string s;
  std::vector<char> buf(32);
  for (auto i : vec) {
    if (!s.empty()) {
      s.push_back(',');
    }
    snprintf(buf.data(), buf.size(), "%d", i);
    s.append(buf.data());
  }
  return s;
}

std::string vector_str(const std::vector<char> &vec) {
  std::string s;
  std::vector<char> buf(32);
  for (auto i : vec) {
    if (!s.empty()) {
      s.push_back(',');
    }
    snprintf(buf.data(), buf.size(), "%02x", (unsigned char)i);
    s.append(buf.data());
  }
  return s;
}

bool is_little_endian(const compound_doc_header_t &hdr) {
  return hdr.byte_order == 0xfeff;
}

bool is_big_endian(const compound_doc_header_t &hdr) {
  return hdr.byte_order == 0xfffe;
}

size_t sect_pos(const compound_doc_header_t &hdr, uint32_t sectid) {
  return 512ul + static_cast<size_t>(sectid) * (1ul << hdr.ssz);
}

size_t short_sect_pos(const compound_doc_header_t &hdr, uint32_t sectid) {
  return static_cast<size_t>(sectid) * (1ul << hdr.sssz);
}

// =============================================================================

const uint64_t CompoundDocument::_compound_document_doc_id = 0xE11AB1A1E011CFD0;

int CompoundDocument::ParseFromFile(const std::string &filename) {
  std::vector<char> data;
  utils::read_file(filename.c_str(), &data);
  if (data.size() < sizeof(compound_doc_header_t)) {
    return -1;
  }
  return ParseFromBytes(std::move(data));
}

int CompoundDocument::ParseFromBytes(const char *data, size_t data_len) {
  std::vector<char> vec(data_len);
  memcpy(vec.data(), data, data_len);
  return ParseFromBytes(std::move(vec));
}

int CompoundDocument::get_master_sector_alloc_table(
    const std::vector<char> &data, SectorAllocTable *msat) {
  auto hdr = reinterpret_cast<const compound_doc_header_t *>(data.data());
  for (int i = 0; i < 109; ++i) {
    if (hdr->sec_ids[i] == -1) {
      break;
    }
    msat->push_back(hdr->sec_ids[i]);
  }

  for (int32_t next_sec_id = hdr->sec_id_of_1st_sect_of_master_sect_alloc_table;
       next_sec_id != kEndOfChainSecID;) {
    const char *sec_p;
    size_t size;
    if (get_sector(data, next_sec_id, &sec_p, &size) != 0) {
      return -1;
    }

    size_t cnt = (size - 4) / 4;
    auto ids = reinterpret_cast<const int32_t *>(sec_p);
    for (size_t i = 0; i < cnt; ++i) {
      if (ids[i] != -1) {
        break;
      }
      msat->push_back(ids[i]);
    }
    next_sec_id = ids[cnt];
  }
  return 0;
}

int CompoundDocument::get_sector_alloc_table(const std::vector<char> &data,
                                             const SectorAllocTable &msat,
                                             SectorAllocTable *sat) {
  for (auto sec_id : msat) {
    const char *sec_p;
    size_t size;
    if (get_sector(data, sec_id, &sec_p, &size) != 0) {
      return -1;
    }

    size_t cnt = size / 4;
    auto ids = reinterpret_cast<const int32_t *>(sec_p);
    for (size_t i = 0; i < cnt; ++i) {
      sat->push_back(ids[i]);
    }
  }
  return 0;
}

int CompoundDocument::get_sector(const std::vector<char> &data, int32_t sec_id,
                                 const char **p, size_t *size) {
  auto hdr = reinterpret_cast<const compound_doc_header_t *>(data.data());
  *size = 1ul << hdr->ssz;

  size_t pos = 512ul + static_cast<size_t>(sec_id) * *size;
  if (pos >= data.size()) {
    return -1;
  }
  *p = data.data() + pos;
  if (pos + *size > data.size()) {
    *size = data.size() - pos;
  }
  return 0;
}

int CompoundDocument::get_short_stream_sector(
    const compound_doc_header_t *hdr,
    const std::vector<char> &short_stream_data, int32_t sec_id, const char **p,
    size_t *size) {
  *size = 1ul << hdr->sssz;

  size_t pos = static_cast<size_t>(sec_id) * *size;
  if (*size > short_stream_data.size() ||
      pos + *size > short_stream_data.size()) {
    return -1;
  }
  *p = short_stream_data.data() + pos;
  return 0;
}

int CompoundDocument::get_short_sector_alloc_table(
    const std::vector<char> &data, const SectorAllocTable &sat,
    SectorAllocTable *ssat) {
  auto hdr = reinterpret_cast<const compound_doc_header_t *>(data.data());
  if (hdr->sec_id_of_1st_sect_of_ss_alloc_table < 0) {
    return 0;
  }

  std::vector<int32_t> chain;
  if (get_sec_ids_chain(hdr->sec_id_of_1st_sect_of_ss_alloc_table, sat,
                        &chain) != 0) {
    return -1;
  }

  for (int32_t sec_id : chain) {
    const char *sec_p;
    size_t size;
    if (get_sector(data, sec_id, &sec_p, &size) != 0) {
      return -1;
    }

    size_t cnt = size / 4;
    auto ids = reinterpret_cast<const int32_t *>(sec_p);
    for (size_t i = 0; i < cnt; ++i) {
      ssat->push_back(ids[i]);
    }
  }

  return 0;
}

int CompoundDocument::get_directory_entries(
    const std::vector<char> &data, const SectorAllocTable &sat,
    std::vector<directory_entry_t> *dir_entries) {
  auto hdr = reinterpret_cast<const compound_doc_header_t *>(data.data());
  std::vector<int32_t> chain;
  if (get_sec_ids_chain(hdr->sec_id_of_1st_sect_of_dir_stream, sat, &chain) !=
      0) {
    return -1;
  }

  for (int32_t sec_id : chain) {
    const char *sec_p;
    size_t size;
    if (get_sector(data, sec_id, &sec_p, &size) != 0) {
      return -1;
    }

    size_t cnt = size / sizeof(directory_entry_t);
    auto dirs = reinterpret_cast<const directory_entry_t *>(sec_p);
    for (size_t i = 0; i < cnt; ++i) {
      dir_entries->push_back(dirs[i]);
    }
  }

  return 0;
}

int CompoundDocument::get_short_stream_data(
    const std::vector<char> &data, const SectorAllocTable &ssat,
    const std::vector<directory_entry_t> &dir_entries,
    std::vector<char> *short_stream_data) {
  int32_t first_short_sec_id = kFreeSecID;
  for (auto &dir : dir_entries) {
    if (dir.type == kDirEntryTypeRootStorage) {
      first_short_sec_id = dir.sec_id_of_1st_x;
      break;
    }
  }
  if (first_short_sec_id == kFreeSecID) {
    return -1;
  }
  if (first_short_sec_id == kEndOfChainSecID) {
    return 0;
  }

  std::vector<int32_t> chain;
  if (get_sec_ids_chain(first_short_sec_id, ssat, &chain) != 0) {
    return -1;
  }

  auto hdr = reinterpret_cast<const compound_doc_header_t *>(data.data());
  size_t sec_size = 1ul << hdr->ssz;
  short_stream_data->resize(sec_size * chain.size());
  char *p = short_stream_data->data();
  for (int32_t sec_id : chain) {
    const char *sec_p;
    size_t size;
    if (get_sector(data, sec_id, &sec_p, &size) != 0) {
      return -1;
    }
    memcpy(p, sec_p, size);
    p += size;
  }

  return 0;
}

int CompoundDocument::get_stream(int32_t first_sec_id,
                                 const SectorAllocTable &xsat,
                                 get_sec_ids_t func,
                                 std::vector<char> *stream) const {
  std::vector<int32_t> chain;
  if (get_sec_ids_chain(first_sec_id, xsat, &chain) != 0) {
    return -1;
  }

  char *p = stream->data();
  size_t plen = stream->size();

  for (auto sec_id : chain) {
    const char *sec_p;
    size_t size;
    if ((this->*func)(sec_id, &sec_p, &size) != 0) {
      return -1;
    }

    size_t len = std::min(size, plen);
    memcpy(p, sec_p, len);
    p += len;
    plen -= len;
  }
  return 0;
}

int CompoundDocument::get_sec_ids_chain(int32_t first_xsec_id,
                                        const SectorAllocTable &xsat,
                                        std::vector<int32_t> *chain) {
  for (int32_t sec_id = first_xsec_id; sec_id != kEndOfChainSecID;
       sec_id = xsat[chain->back()]) {
    if (sec_id < 0 || sec_id >= xsat.size()) {
      return -1;
    }

    chain->push_back(sec_id);
    if (chain->size() > xsat.size()) {
      return -1;
    }
  }
  return 0;
}

int CompoundDocument::GetDirEntryStream(const directory_entry_t &dir_entry,
                                        std::vector<char> *stream) const {
  if (dir_entry.sec_id_of_1st_x < 0) {
    return -1;
  }

  auto hdr = reinterpret_cast<const compound_doc_header_t *>(m_data.data());
  stream->resize(dir_entry.size_of_x);
  return get_stream(
      dir_entry.sec_id_of_1st_x,
      dir_entry.size_of_x < hdr->min_size_of_std_stream ? m_ssat : m_sat,
      dir_entry.size_of_x < hdr->min_size_of_std_stream
          ? &CompoundDocument::GetShortStreamSector
          : &CompoundDocument::GetSector,
      stream);
}

// =============================================================================

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
    slog(Err, "u16_to_u8 err, %s", e.what());
    return -1;
  }
}

}  // namespace msoffice
