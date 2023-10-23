#pragma once

#include <stdint.h>

#include <limits>
#include <map>
#include <string>
#include <vector>

#define _STYLE_Info  "\e[3;32m"
#define _STYLE_Err   "\e[3;31m"
#define _STYLE_Warn  "\e[3;33m"
#define _STYLE_Debug "\e[3;36m"
#define xmlog(_type_, _fmt_, ...)                                     \
  printf(_STYLE_##_type_ "%.1s [%s:%s:%d]\e[0m " _fmt_ "\n", #_type_, \
         __FILE__, __func__, __LINE__, ##__VA_ARGS__)

namespace msoffice {

template <typename T>
struct is_bytes : public std::integral_constant<
                      bool, std::is_same<std::vector<char>,
                                         typename std::decay<T>::type>::value> {
};

std::string vector_str(const std::vector<int>& vec);

struct fetch_text_options_t {
  size_t max_fetch_text_len = std::numeric_limits<size_t>::max();
  bool fetch_text_from_drawing = false;
  std::string xls_delimiter = ",";
  bool xls_skip_blank_cell = true;
};

enum spec_sect_id_t {
  kFreeSecID = -1,
  kEndOfChainSecID = -2,
  kSatSecID = -3,
  kMsatSecID = -4,
};

struct compound_doc_header_t {
  uint64_t doc_id;
  uint8_t uid[16];
  uint16_t revision;
  uint16_t version;
  uint16_t byte_order;
  uint16_t ssz;
  uint16_t sssz;
  uint8_t not_used0[10];
  uint32_t total_number_of_sect_used_for_sect_alloc_table;
  int32_t sec_id_of_1st_sect_of_dir_stream;
  uint8_t not_used1[4];
  uint32_t min_size_of_std_stream;
  int32_t sec_id_of_1st_sect_of_ss_alloc_table;
  uint32_t total_number_of_sect_used_for_ss_alloc_table;
  int32_t sec_id_of_1st_sect_of_master_sect_alloc_table;
  uint32_t total_number_of_sect_used_for_master_sect_alloc_table;
  int32_t sec_ids[109];
} __attribute__((packed));

enum directory_entry_type_t {
  kDirEntryTypeEmpty = 0x00,
  kDirEntryTypeUserStorage = 0x01,
  kDirEntryTypeUserStream = 0x02,
  kDirEntryTypeLockBytes = 0x03,
  kDirEntryTypeProperty = 0x04,
  kDirEntryTypeRootStorage = 0x05,
};

enum directory_entry_color_t {
  kDirEntryColorRed = 0x00,
  kDirEntryColorBlack = 0x01,
};

struct directory_entry_t {
  char16_t unicode_name[32];
  uint16_t name_len;
  uint8_t type;
  uint8_t color;
  int32_t left_child_dir_id;
  int32_t right_child_dir_id;
  int32_t root_dir_id;
  char uid[16];
  uint32_t user_flags;
  uint64_t ctime;
  uint64_t mtime;
  int32_t sec_id_of_1st_x;
  int32_t size_of_x;
  char unused[4];
} __attribute__((packed));

class CompoundDocument {
 public:
  using SectorAllocTable = std::vector<int32_t>;

  using StreamSecIdChain = std::vector<int32_t>;
  using StreamSecIdChainTable = std::map<int32_t, StreamSecIdChain>;

  int ParseFromFile(const std::string& filename);
  int ParseFromBytes(const char* data, size_t data_len);

  template <typename _T,
            typename std::enable_if<is_bytes<_T>::value>::type* = nullptr>
  int ParseFromBytes(_T&& data) {
    auto hdr = reinterpret_cast<compound_doc_header_t*>(data.data());
    if (hdr->doc_id != _compound_document_doc_id) {
      return -1;
    }

    if (get_master_sector_alloc_table(data, &this->m_msat) != 0) {
      return -1;
    }

    if (get_sector_alloc_table(data, this->m_msat, &this->m_sat) != 0) {
      return -1;
    }

    if (get_stream_sec_id_chains(this->m_sat, &this->m_stream_sec_id_chains) !=
        0) {
      return -1;
    }

    if (get_short_sector_alloc_table(data, this->m_stream_sec_id_chains,
                                     &this->m_ssat) != 0) {
      return -1;
    }

    if (get_stream_sec_id_chains(this->m_ssat,
                                 &this->m_short_stream_sec_id_chains) != 0) {
      return -1;
    }

    if (get_directory_entries(data, this->m_stream_sec_id_chains,
                              &this->m_dir_entries) != 0) {
      return -1;
    }

    if (get_short_stream_data(data, this->m_stream_sec_id_chains,
                              this->m_dir_entries,
                              &this->m_short_stream_data) != 0) {
      return -1;
    }

    this->m_data = std::forward<_T>(data);
    return 0;
  }

  inline const compound_doc_header_t* GetHeader() const {
    return reinterpret_cast<const compound_doc_header_t*>(m_data.data());
  }

  inline const SectorAllocTable& GetMSAT() const { return m_msat; }

  inline const SectorAllocTable& GetSAT() const { return m_sat; }

  inline const SectorAllocTable& GetSSAT() const { return m_ssat; }

  inline const StreamSecIdChainTable& GetStreamSecIdChain() const {
    return m_stream_sec_id_chains;
  }

  inline const StreamSecIdChainTable& GetShortStreamSecIdChain() const {
    return m_short_stream_sec_id_chains;
  }

  inline const std::vector<directory_entry_t> GetDirEntries() const {
    return m_dir_entries;
  }

  inline int GetSector(int32_t sec_id, const char** p, size_t* size) const {
    return get_sector(m_data, sec_id, p, size);
  }

  inline int GetShortStreamSector(int32_t sec_id, const char** p,
                                  size_t* size) const {
    auto hdr = reinterpret_cast<const compound_doc_header_t*>(m_data.data());
    return get_short_stream_sector(hdr, m_short_stream_data, sec_id, p, size);
  }

  int GetDirEntryStream(const directory_entry_t& dir_entry,
                        std::vector<char>* stream) const;

 private:
  static const uint64_t _compound_document_doc_id;

  using get_sec_ids_t = int (CompoundDocument::*)(int32_t, const char**,
                                                  size_t*) const;

  int get_stream(int32_t first_sec_id, const StreamSecIdChainTable& tab,
                 get_sec_ids_t func, std::vector<char>* stream) const;

  static int get_master_sector_alloc_table(const std::vector<char>& data,
                                           SectorAllocTable* msat);

  static int get_sector_alloc_table(const std::vector<char>& data,
                                    const SectorAllocTable& msat,
                                    SectorAllocTable* sat);

  static int get_short_sector_alloc_table(
      const std::vector<char>& data,
      const StreamSecIdChainTable& stream_sec_id_chains,
      SectorAllocTable* ssat);

  static int get_sector(const std::vector<char>& data, int32_t sec_id,
                        const char** p, size_t* size);

  static int get_short_stream_sector(const compound_doc_header_t* hdr,
                                     const std::vector<char>& short_stream_data,
                                     int32_t sec_id, const char** p,
                                     size_t* size);

  static int get_stream_sec_id_chains(
      const SectorAllocTable& sat, StreamSecIdChainTable* strem_sec_id_chains);

  static int get_directory_entries(
      const std::vector<char>& data,
      const StreamSecIdChainTable& stream_sec_id_chains,
      std::vector<directory_entry_t>* dir_entries);

  static int get_short_stream_data(
      const std::vector<char>& data,
      const StreamSecIdChainTable& stream_sec_id_chains,
      const std::vector<directory_entry_t>& dir_entries,
      std::vector<char>* short_stream_data);

 private:
  std::vector<char> m_data;
  std::vector<char> m_short_stream_data;
  SectorAllocTable m_msat;
  SectorAllocTable m_sat;
  SectorAllocTable m_ssat;
  StreamSecIdChainTable m_stream_sec_id_chains;
  StreamSecIdChainTable m_short_stream_sec_id_chains;
  std::vector<directory_entry_t> m_dir_entries;
};

template <typename T>
struct is_compound_document
    : public std::integral_constant<
          bool,
          std::is_same<CompoundDocument, typename std::decay<T>::type>::value> {
};

void RemoveControlCharacter(std::string* text);

int Utf16ToUtf8(const char16_t* begin, const char16_t* end, std::string* u8);

}  // namespace msoffice

#undef xmlog
