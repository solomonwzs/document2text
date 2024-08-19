#include "msoffice/compound_document.h"

#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include <cmath>

#include "utils/utils.h"

#define _STYLE_Info "\e[3;32m"
#define _STYLE_Err "\e[3;31m"
#define _STYLE_Warn "\e[3;33m"
#define _STYLE_Debug "\e[3;36m"
#define slog(_type_, _fmt_, ...)                                      \
  printf(_STYLE_##_type_ "%.1s [%s:%s:%d]\e[0m " _fmt_ "\n", #_type_, \
         __FILE__, __func__, __LINE__, ##__VA_ARGS__)

#define CONCAT_(_A, _B) _A##_B
#define CONCAT(_A, _B) CONCAT_(_A, _B)
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

void rc4_setup(rc4_state_t *state, const uint8_t *key, size_t klen) {
  state->x = 0;
  state->y = 0;
  state->m.resize(256);

  for (int i = 0; i < 256; ++i) {
    state->m[i] = i;
  }

  for (int i = 0, j = 0; i < 256; ++i) {
    j = static_cast<uint8_t>((j + state->m[i] + key[i % klen]) & 0xff);
    uint8_t tmp = state->m[i];
    state->m[i] = state->m[j];
    state->m[j] = tmp;
  }
}

std::vector<uint8_t> rc4_crypt(rc4_state_t *state, const uint8_t *data,
                               size_t dlen) {
  uint8_t x = state->x;
  uint8_t y = state->y;

  std::vector<uint8_t> out(dlen);
  memcpy(out.data(), data, dlen);
  for (size_t i = 0; i < dlen; ++i) {
    x = x + 1;
    y = y + state->m[i];

    uint8_t tmp = x;
    x = y;
    y = tmp;

    out[i] ^= state->m[(state->m[x] + state->m[y]) & 0xff];
  }
  state->x = x;
  state->y = y;

  return out;
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

  int sec_cnt = std::ceil(static_cast<double>(data.size()) / (1 << hdr->ssz));
  for (int32_t next_sec_id = hdr->sec_id_of_1st_sect_of_master_sect_alloc_table,
               n = 0;
       next_sec_id != kEndOfChainSecID &&
       n < static_cast<int32_t>(
               hdr->total_number_of_sect_used_for_master_sect_alloc_table) &&
       n < sec_cnt;
       ++n) {
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
    if (sec_id < 0 || sec_id >= static_cast<int>(xsat.size())) {
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
      dir_entry.size_of_x < static_cast<int>(hdr->min_size_of_std_stream)
          ? m_ssat
          : m_sat,
      dir_entry.size_of_x < static_cast<int>(hdr->min_size_of_std_stream)
          ? &CompoundDocument::GetShortStreamSector
          : &CompoundDocument::GetSector,
      stream);
}

std::set<int> CompoundDocument::GetValidDirIndex() const {
  std::set<int> res;
  // if (m_dir_entries.empty()) {
  //   return res;
  // }

  // std::vector<int> idx;
  // idx.reserve(m_dir_entries.size());
  // idx.push_back(0);

  // for (size_t i = 0; i < idx.size(); ++i) {
  //   res.insert(idx[i]);
  //   auto &dir = m_dir_entries[idx[i]];
  //   xmlog(Debug, "%d, %d", idx[i], m_sat[dir.sec_id_of_1st_x]);

  //   if (dir.root_dir_id != -1 && dir.root_dir_id < m_dir_entries.size() &&
  //       res.find(dir.root_dir_id) == res.end()) {
  //     idx.push_back(dir.root_dir_id);
  //   }
  //   if (dir.left_child_dir_id != -1 &&
  //       dir.left_child_dir_id < m_dir_entries.size() &&
  //       res.find(dir.left_child_dir_id) == res.end()) {
  //     idx.push_back(dir.left_child_dir_id);
  //   }
  //   if (dir.right_child_dir_id != -1 &&
  //       dir.right_child_dir_id < m_dir_entries.size() &&
  //       res.find(dir.right_child_dir_id) == res.end()) {
  //     idx.push_back(dir.right_child_dir_id);
  //   }
  // }
  return res;
}

}  // namespace msoffice
