#include "msoffice/ms_ppt.h"

#include <stdio.h>
#include <string.h>

#include <functional>

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

namespace ppt {

static std::string g_PowerPointDocDirName = "PowerPoint Document";
static std::string g_CurrentUserDirName = "Current User";

ssize_t CurrentUserAtom::ReadAndParse(const char *data, size_t size) {
  if (size < sizeof(hdr_t)) {
    return -1;
  }

  auto hdr = reinterpret_cast<const hdr_t *>(data);
  if (hdr->rh.recVer() != 0 || hdr->rh.recInstance() != 0 ||
      hdr->rh.recType != kRT_CurrentUserAtom || hdr->lenUserName > 255 ||
      hdr->docFileVersion != 0x03F4 || hdr->majorVersion != 0x03 ||
      hdr->minorVersion != 0x00 ||
      (hdr->headerToken != 0xE391C05F && hdr->headerToken != 0xF3D1C4DF)) {
    return -1;
  }

  const size_t hdr_size = sizeof(hdr_t) + hdr->lenUserName + sizeof(uint32_t) +
                          2 * hdr->lenUserName;
  if (size < hdr_size) {
    return -1;
  }

  auto relVersion = reinterpret_cast<const uint32_t *>(data + sizeof(hdr_t) +
                                                       hdr->lenUserName);
  if (*relVersion != 0x00000008 && *relVersion != 0x00000009) {
    return -1;
  }

  auto unicode_name = reinterpret_cast<const char16_t *>(
      data + sizeof(hdr_t) + hdr->lenUserName + sizeof(uint32_t));
  std::string user_name;
  if (Utf16ToUtf8(unicode_name, unicode_name + hdr->lenUserName, &user_name) !=
      0) {
    return -1;
  }

  m_hdr = *hdr;
  m_rel_version = *relVersion;
  m_ansi_user_name = std::string(data + sizeof(hdr_t), hdr->lenUserName);
  m_user_name = std::move(user_name);

  return hdr_size;
}

// =============================================================================

int PersistDirectoryAtom::ParsePersistDirectoryEntry(
    const UserEditAtom_t &user_edit_atom, const char *data, size_t data_len,
    std::vector<const PersistDirectoryEntry_t *> *entries) {
  entries->clear();
  for (size_t offset = 0; offset < data_len;) {
    auto entry =
        reinterpret_cast<const PersistDirectoryEntry_t *>(data + offset);
    size_t end = offset + sizeof(PersistDirectoryEntry_t) +
                 entry->cPersist() * sizeof(uint32_t);
    if (end > data_len) {
      return -1;
    }

    if (entry->cPersist() == 0) {
      return -1;
    }

    for (uint16_t i = 0; i < entry->cPersist(); ++i) {
      if (entry->persistOffset[i] < user_edit_atom.offsetLastEdit) {
        return -1;
      }
    }

    entries->push_back(entry);
    offset +=
        sizeof(PersistDirectoryEntry_t) + entry->cPersist() * sizeof(uint32_t);
  }
  return 0;
}

ssize_t PersistDirectoryAtom::ReadAndParse(const UserEditAtom_t &user_edit_atom,
                                           const char *data, size_t len) {
  if (len < sizeof(record_header_t)) {
    return -1;
  }
  auto rh = reinterpret_cast<const record_header_t *>(data);
  if (rh->recVer() != 0 || rh->recInstance() != 0 ||
      rh->recType != kRT_PersistDirectoryAtom ||
      len < sizeof(record_header_t) + rh->recLen) {
    return -1;
  }
  m_rh = *rh;

  m_data.resize(rh->recLen);
  memcpy(m_data.data(), data + sizeof(record_header_t), rh->recLen);
  if (ParsePersistDirectoryEntry(user_edit_atom, m_data.data(), m_data.size(),
                                 &m_persistDirectoryEntry) != 0) {
    return -1;
  }

  return sizeof(record_header_t) + rh->recLen;
}

// =============================================================================

using fetch_text_func_t = std::function<int(
    const char *, size_t, fetch_text_options_t &, std::string *)>;

static int atom_list_fetch_text(const char *container_data,
                                size_t container_len,
                                fetch_text_options_t &opts, std::string *text);

static int fetch_text_TextBytesAtom(const char *container_data,
                                    size_t container_len,
                                    fetch_text_options_t &opts,
                                    std::string *text) {
  size_t len = std::min(container_len, opts.max_fetch_text_len);
  text->append(container_data, len);
  opts.max_fetch_text_len -= len;
  return 0;
}

static int fetch_text_TextCharsAtom(const char *container_data,
                                    size_t container_len,
                                    fetch_text_options_t &opts,
                                    std::string *text) {
  size_t len = std::min(container_len / 2, opts.max_fetch_text_len);
  auto begin = reinterpret_cast<const char16_t *>(container_data);
  auto end = reinterpret_cast<const char16_t *>(container_data + len * 2);

  std::string s;
  if (Utf16ToUtf8(begin, end, &s) != 0) {
    return -1;
  }
  text->append(std::move(s));
  opts.max_fetch_text_len -= len;
  return 0;
}

// =============================================================================

static int atom_list_fetch_text(const char *container_data,
                                size_t container_len,
                                fetch_text_options_t &opts, std::string *text) {
  const char *end = container_data + container_len;
  size_t i;
  uint32_t step;
  for (i = 0; i < container_len && opts.max_fetch_text_len > 0; i += step) {
    const char *p = container_data + i;
    if (p >= end || p + sizeof(record_header_t) > end) {
      return -1;
    }
    auto rh = reinterpret_cast<const record_header_t *>(p);
    step = sizeof(record_header_t) + rh->recLen;

    fetch_text_func_t fetch_text_func = nullptr;
    if (rh->recType == kRT_SlideListWithText) {
      fetch_text_func = atom_list_fetch_text;
    } else if (opts.fetch_text_from_drawing &&
               (rh->recType == kRT_Drawing || rh->recType == kRT_OfficeArtDg ||
                rh->recType == kRT_OfficeArtSpgrContainer ||
                rh->recType == kRT_OfficeArtSpContainer ||
                rh->recType == kRT_OfficeArtClientTextbox)) {
      fetch_text_func = atom_list_fetch_text;
    } else if (rh->recType == kRT_TextBytesAtom) {
      fetch_text_func = fetch_text_TextBytesAtom;
    } else if (rh->recType == kRT_TextCharsAtom) {
      fetch_text_func = fetch_text_TextCharsAtom;
    }

    if (fetch_text_func != nullptr) {
      if (fetch_text_func(p + sizeof(record_header_t), rh->recLen, opts,
                          text) != 0) {
        return -1;
      } else {
        if (opts.max_fetch_text_len > 0 && !text->empty() &&
            text->back() != '\n') {
          text->push_back('\n');
          opts.max_fetch_text_len -= 1;
        }
      }
    }
  }

  return 0;
}

// =============================================================================

int MsPPT::ParseFromFile(const std::string &filename) {
  std::vector<char> data;
  utils::read_file(filename.c_str(), &data);
  if (data.size() < sizeof(compound_doc_header_t)) {
    return -1;
  }
  return ParseFromBytes(std::move<>(data));
}

int MsPPT::parse() {
  int32_t idx_current_user = -1;
  int32_t idx_ppt_doc = -1;

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

    if (dirname == g_CurrentUserDirName) {
      idx_current_user = i;
    } else if (dirname == g_PowerPointDocDirName) {
      idx_ppt_doc = i;
    }
  }
  if (idx_ppt_doc == -1 || idx_current_user == -1) {
    return -1;
  }

  m_idx_current_user = idx_current_user;
  m_idx_ppt_doc = idx_ppt_doc;
  return 0;
}

int MsPPT::GetPersistId2Offset(const CurrentUserAtom &current_user_atom,
                               const char *ppt_doc_stream,
                               size_t ppt_doc_stream_len,
                               std::map<uint32_t, uint32_t> *id2offset) {
  const char *end = ppt_doc_stream + ppt_doc_stream_len;
  const char *p;
  for (size_t offset = current_user_atom.Header().offsetToCurrentEdit;;) {
    p = ppt_doc_stream + offset;
    if (p >= end || p + sizeof(UserEditAtom_t) > end) {
      return -1;
    }
    auto user_edit_atom = reinterpret_cast<const UserEditAtom_t *>(p);
    if (!(user_edit_atom->rh.recVer() == 0x00 &&
          user_edit_atom->rh.recInstance() == 0x000 &&
          user_edit_atom->rh.recType == kRT_UserEditAtom &&
          (user_edit_atom->rh.recLen == 0x1C ||
           user_edit_atom->rh.recLen == 0x20) &&
          user_edit_atom->minorVersion == 0x00 &&
          user_edit_atom->majorVersion == 0x03 &&
          user_edit_atom->docPersistIdRef == 0x00000001)) {
      return -1;
    }

    p = ppt_doc_stream + user_edit_atom->offsetPersistDirectory;
    if (p >= end || p + sizeof(record_header_t) > end) {
      return -1;
    }
    auto rh = reinterpret_cast<const record_header_t *>(p);
    if (rh->recVer() != 0 || rh->recInstance() != 0 ||
        rh->recType != kRT_PersistDirectoryAtom ||
        p + sizeof(record_header_t) + rh->recLen > end) {
      return -1;
    }

    p += sizeof(record_header_t);
    std::vector<const PersistDirectoryEntry_t *> entries;
    if (PersistDirectoryAtom::ParsePersistDirectoryEntry(
            *user_edit_atom, p, rh->recLen, &entries) != 0) {
      return -1;
    }

    for (auto i : entries) {
      uint32_t id = i->persistId();
      for (int j = 0; j < i->cPersist(); ++j) {
        if (id2offset->find(id + j) != id2offset->end()) {
          continue;
        }
        (*id2offset)[id + j] = i->persistOffset[j];
      }
    }

    offset = user_edit_atom->offsetLastEdit;
    if (offset == 0) {
      break;
    }
  }
  return 0;
}

int MsPPT::FetchText(const fetch_text_options_t *user_opts,
                     std::string *text) const {
  fetch_text_options_t opts =
      user_opts != nullptr ? *user_opts : __defaultFetchTextOptions;

  auto &dirs = m_comp_doc.GetDirEntries();

  std::vector<char> current_user_stream;
  if (m_comp_doc.GetDirEntryStream(dirs[m_idx_current_user],
                                   &current_user_stream) != 0 ||
      current_user_stream.size() < sizeof(record_header_t)) {
    return -1;
  }
  auto rh = reinterpret_cast<record_header_t *>(current_user_stream.data());
  if (rh->recType != kRT_CurrentUserAtom) {
    return -1;
  }

  CurrentUserAtom current_user_atom;
  ssize_t current_user_atom_len = current_user_atom.ReadAndParse(
      current_user_stream.data(), current_user_stream.size());
  if (current_user_atom_len < 0) {
    return -1;
  }
  if (current_user_atom.IsEncrypted()) {
    return -1;
  }

  std::vector<char> ppt_doc_stream;
  if (m_comp_doc.GetDirEntryStream(dirs[m_idx_ppt_doc], &ppt_doc_stream) != 0) {
    return -1;
  }

  std::map<uint32_t, uint32_t> id2offset;
  if (GetPersistId2Offset(current_user_atom, ppt_doc_stream.data(),
                          ppt_doc_stream.size(), &id2offset) != 0) {
    return -1;
  }

  const char *end = ppt_doc_stream.data() + ppt_doc_stream.size();
  for (auto i : id2offset) {
    const char *p = ppt_doc_stream.data() + i.second;
    if (p >= end || p + sizeof(record_header_t) > end) {
      return -1;
    }
    auto rh = reinterpret_cast<const record_header_t *>(p);

    fetch_text_func_t fetch_text_func = nullptr;
    if (rh->recType == kRT_Document) {
      fetch_text_func = atom_list_fetch_text;
    } else if (rh->recType == kRT_Slide) {
      fetch_text_func = atom_list_fetch_text;
    } else {
      continue;
    }

    p += sizeof(record_header_t);
    if (p >= end || p + rh->recLen > end) {
      return -1;
    }

    if (fetch_text_func(p, rh->recLen, opts, text) != 0) {
      return -1;
    } else if (opts.max_fetch_text_len == 0) {
      break;
    }
  }

  RemoveControlCharacter(text);

  return 0;
}

}  // namespace ppt

}  // namespace msoffice
