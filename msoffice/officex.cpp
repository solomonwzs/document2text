#include "msoffice/officex.h"

#include <stdio.h>
#include <string.h>
#include <zip.h>

#include <limits>
#include <memory>

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

namespace officex {

ZipHelper::ZipHelper() : m_zsrc(nullptr), m_zfd(nullptr) {}

ZipHelper::~ZipHelper() {
  if (m_zfd != nullptr) {
    zip_discard(m_zfd);
    m_zfd = nullptr;
  }

  if (m_zsrc != nullptr) {
    zip_source_close(m_zsrc);
    zip_source_free(m_zsrc);
    m_zsrc = nullptr;
  }
}

int ZipHelper::OpenFromBytes(const std::vector<char> &data) {
  return OpenFromBytes(data.data(), data.size());
}

int ZipHelper::OpenFromBytes(const char *data, size_t data_len) {
  if (data_len < sizeof(uint32_t) ||
      *reinterpret_cast<const uint32_t *>(data) != 0x04034b50) {
    return -1;
  }

  zip_error_t zerr;
  zip_source_t *zsrc = nullptr;
  zip_t *zfd = nullptr;
  std::map<std::string, int64_t> name2idx;

  zip_error_init(&zerr);

  zsrc = zip_source_buffer_create(data, data_len, 0, &zerr);
  if (zsrc == nullptr) {
    slog(Err, "%s", zip_error_strerror(&zerr));
    zip_error_fini(&zerr);
    return -1;
  }

  if (zip_source_open(zsrc) != 0) {
    slog(Err, "zip_source_open fail");
    goto _ERR_HAS_BUF_CREATE;
  }

  zip_source_keep(zsrc);

  zfd = zip_open_from_source(zsrc, ZIP_RDONLY, &zerr);
  if (zfd == nullptr) {
    slog(Err, "%s", zip_error_strerror(&zerr));
    zip_error_fini(&zerr);
    goto _ERR_HAS_SRC_OPEN;
  }

  for (int64_t i = 0; i < zip_get_num_entries(zfd, 0); ++i) {
    const char *name = zip_get_name(zfd, i, 0);
    name2idx.insert({name, i});
  }
  goto _INIT_OK;

_ERR_HAS_SRC_OPEN:
  zip_source_close(zsrc);
_ERR_HAS_BUF_CREATE:
  zip_source_free(zsrc);
  return -1;

_INIT_OK:
  m_zsrc = zsrc;
  m_zfd = zfd;
  m_name2idx.swap(name2idx);
  return 0;
}

template <typename T>
static int zip_read_by_name(const std::map<std::string, int64_t> &name2idx,
                            zip_t *zfd, const std::string &name,
                            size_t max_read_len, T *data) {
  auto it = name2idx.find(name);
  if (it == name2idx.end()) {
    return -1;
  }

  zip_stat_t zs;
  zip_stat_init(&zs);
  if (zip_stat_index(zfd, it->second, 0, &zs) != 0) {
    return -1;
  }

  zip_file_t *zfile = zip_fopen_index(zfd, it->second, 0);
  if (zfile == nullptr) {
    return -1;
  }
  _defer([&](...) { zip_fclose(zfile); });

  data->resize(std::min(zs.size, max_read_len));
  if (zip_fread(zfile, &((*data)[0]), data->size()) < 0) {
    return -1;
  }
  return 0;
}

int ZipHelper::ReadByName(const std::string &name, size_t max_read_len,
                          std::vector<char> *data) {
  return zip_read_by_name(m_name2idx, m_zfd, name, max_read_len, data);
}

int ZipHelper::ReadByName(const std::string &name, size_t max_read_len,
                          std::string *data) {
  return zip_read_by_name(m_name2idx, m_zfd, name, max_read_len, data);
}

// =============================================================================

class _TempTruncateStr {
 public:
  _TempTruncateStr(const char *p, size_t len)
      : m_ch(p[len]), m_pos(const_cast<char *>(p) + len) {
    *m_pos = 0;
  }
  ~_TempTruncateStr() {
    *m_pos = m_ch;
  }

 private:
  char m_ch;
  char *m_pos;
};

#define _temp_truncate_string(_p_, _len_) \
  _TempTruncateStr CONCAT(__tmp_truncate_str, __LINE__)(_p_, _len_)

static inline bool is_sst_tag(const char *tag, size_t tag_len) {
  _temp_truncate_string(tag, tag_len);
  const char *p = strstr(tag, "t=\"s\"");
  return p != nullptr;
}

static bool get_attr_value(const char *tag, size_t tag_len, const char *name,
                           std::string *val) {
  _temp_truncate_string(tag, tag_len);
  const char *p = strstr(tag, name);
  if (p == nullptr) {
    return false;
  }

  size_t nlen = strlen(name);
  const char *quot = p + nlen;
  if (*quot != '"' && *quot != '\'') {
    return false;
  }

  const char *i = quot + 1;
  for (; i < tag + tag_len && *i != *quot; ++i) {
  }
  if (i == tag + tag_len) {
    return false;
  }

  val->assign(quot + 1, i - (quot + 1));
  return true;
}

static bool is_tag(const char *s, const char *name) {
  size_t nlen = strlen(name);
  if (strncmp(s, name, nlen) != 0) {
    return false;
  }
  const char *p = s + nlen;
  if (*p == '\0' || !(*p == ' ' || *p == '\t' || *p == '\n' || *p == '\r' ||
                      *p == '>' || *p == '/')) {
    return false;
  }
  return true;
}

static inline bool is_empty_element_tag(const char *s, size_t len) {
  return len > 2 && s[len - 2] == '/' && s[len - 1] == '>';
}

static bool find_tag(const char *s, const char *name, const char **out,
                     size_t *out_len) {
  for (const char *left = s; *left != '\0'; ++left) {
    left = strstr(left, name);
    if (left == nullptr) {
      return false;
    }

    const char *right = left + strlen(name);
    if (*right == '\0') {
      return false;
    } else if (!(*right == ' ' || *right == '\t' || *right == '\n' ||
                 *right == '\r' || *right == '>' || *right == '/')) {
      continue;
    }
    if ((right = strchr(right, '>')) == nullptr) {
      return false;
    }

    *out = left;
    *out_len = right - left + 1;
    return true;
  }
  return false;
}

static void append_text(std::string *text, size_t *max_len, const char *s,
                        size_t len, size_t wcnt) {
  if (wcnt <= *max_len) {
    text->append(s, len);
    *max_len -= wcnt;
  } else {
    size_t offset = utils::fix_utf8_word_cnt(s, *max_len);
    text->append(s, offset);
    *max_len = 0;
  }
}

static inline void append_text(std::string *text, size_t *max_len,
                               const char *s, size_t len) {
  size_t wcnt = utils::count_utf8_word_cnt(s, len);
  return append_text(text, max_len, s, len, wcnt);
}

// =============================================================================

int MsDOCxFetchText(char *xml_text, const fetch_text_options_t *opts,
                    std::string *text) {
  if (opts == nullptr) {
    opts = &__defaultFetchTextOptions;
  }
  size_t max_len = opts->max_fetch_text_len;

  const char *ll = nullptr;
  size_t llen = 0;
  for (bool f = find_tag(xml_text, "<w:t", &ll, &llen); f && max_len > 0;
       f = find_tag(ll + llen, "<w:t", &ll, &llen)) {
    const char *lr = ll + llen;
    const char *rl = nullptr;
    size_t rlen = 0;
    if (!find_tag(lr, "</w:t", &rl, &rlen)) {
      break;
    }

    append_text(text, &max_len, lr, rl - lr);
    if (max_len > 0) {
      text->push_back('\n');
      max_len -= 1;
    }
  }

  return 0;
}

int MsDOCxFetchText(ZipHelper &zip, const fetch_text_options_t *opts,
                    std::string *text) {
  if (opts == nullptr) {
    opts = &__defaultFetchTextOptions;
  }
  std::string xml_text;
  if (zip.ReadByName("word/document.xml", opts->xml_max_file_len, &xml_text) !=
      0) {
    return -1;
  }
  return MsDOCxFetchText(&xml_text[0], opts, text);
}

// =============================================================================

int MsPPTxFetchText(char *xml_text, const fetch_text_options_t *opts,
                    std::string *text, size_t *fetch_len) {
  if (opts == nullptr) {
    opts = &__defaultFetchTextOptions;
  }

  if (fetch_len != nullptr) {
    *fetch_len = opts->max_fetch_text_len;
  }
  size_t max_len = opts->max_fetch_text_len;

  const char *ap = nullptr;
  size_t ap_len = 0;
  for (bool f = find_tag(xml_text, "<a:p", &ap, &ap_len); f && max_len > 0;
       f = find_tag(ap + ap_len, "<a:p", &ap, &ap_len)) {
    const char *end_ap;
    size_t end_ap_len = 0;
    if (!find_tag(ap + ap_len, "</a:p", &end_ap, &end_ap_len)) {
      break;
    }
    _temp_truncate_string(ap, end_ap - ap);

    bool has_text = false;
    const char *curr = ap + ap_len;
    for (;;) {
      const char *p = strstr(curr, "<a:");
      if (p == nullptr) {
        break;
      }
      curr = p + 1;

      if (is_tag(p, "<a:br")) {
        has_text = true;

        text->push_back('\n');
        max_len -= 1;
      } else if (is_tag(p, "<a:t")) {
        const char *q = strchr(p, '>');
        if (q != nullptr) {
          q += 1;
          const char *at;
          size_t at_len = 0;
          if (find_tag(q, "</a:t", &at, &at_len)) {
            has_text = true;

            append_text(text, &max_len, q, at - q);
          }
        }
      }
    }

    if (has_text && max_len > 0) {
      text->push_back('\n');
      max_len -= 1;
    }
  }

  if (fetch_len != nullptr) {
    *fetch_len -= max_len;
  }

  return 0;
}

int MsPPTxFetchText(ZipHelper &zip, const fetch_text_options_t *_opts,
                    std::string *text) {
  fetch_text_options_t opts =
      _opts == nullptr ? __defaultFetchTextOptions : *_opts;

  std::vector<char> name(128);
  std::string xml;
  for (size_t i = 1; opts.max_fetch_text_len > 0; ++i) {
    snprintf(name.data(), name.size(), "ppt/slides/slide%ld.xml", i);
    if (zip.ReadByName(name.data(), opts.xml_max_file_len, &xml) != 0) {
      break;
    }

    size_t fetch_len;
    if (MsPPTxFetchText(&xml[0], &opts, text, &fetch_len) != 0) {
      return -1;
    }
    opts.max_fetch_text_len -= fetch_len;
  }
  return 0;
}

// =============================================================================

struct xlsx_tags_t {
  const char *sheets;
  const char *sheets_;
  const char *sheet;
  const char *si;
  const char *si_;
  const char *t;
  const char *t_;
  const char *row;
  const char *row_;
  const char *c;
  const char *c_;
  const char *v;
  const char *v_;
};

static const xlsx_tags_t g_xlsxTags{
    .sheets = "<sheets",
    .sheets_ = "</sheets",
    .sheet = "<sheet",
    .si = "<si",
    .si_ = "</si",
    .t = "<t",
    .t_ = "</t",
    .row = "<row",
    .row_ = "</row",
    .c = "<c",
    .c_ = "</c",
    .v = "<v",
    .v_ = "</v",
};

static const xlsx_tags_t g_xlsxXTags{
    .sheets = "<x:sheets",
    .sheets_ = "</x:sheets",
    .sheet = "<x:sheet",
    .si = "<x:si",
    .si_ = "</x:si",
    .t = "<x:t",
    .t_ = "</x:t",
    .row = "<x:row",
    .row_ = "</x:row",
    .c = "<x:c",
    .c_ = "</x:c",
    .v = "<x:v",
    .v_ = "</x:v",
};

int MsXLSxFetchRelationships(ZipHelper &zip, size_t xml_max_file_len,
                             std::map<std::string, std::string> *rid2target) {
  std::string xml;
  if (zip.ReadByName("xl/_rels/workbook.xml.rels", xml_max_file_len, &xml) !=
      0) {
    return -1;
  }

  const char *rels = nullptr;
  size_t rels_len = 0;
  const char *rels_end = nullptr;
  size_t rels_end_len = 0;
  if (!find_tag(xml.c_str(), "<Relationships", &rels, &rels_len) ||
      !find_tag(rels + rels_len, "</Relationships", &rels_end, &rels_end_len)) {
    return -1;
  }
  _temp_truncate_string(rels, rels_end - rels);

  const char *rel = nullptr;
  size_t rel_len = 0;
  for (bool f = find_tag(rels + rels_len, "<Relationship", &rel, &rel_len); f;
       f = find_tag(rel + rel_len, "<Relationship", &rel, &rel_len)) {
    std::string rid;
    std::string target;
    if (get_attr_value(rel, rel_len, "Id=", &rid) &&
        get_attr_value(rel, rel_len, "Target=", &target) && !rid.empty() &&
        !target.empty()) {
      if (target[0] == '/') {
        target.erase(0, 1);
      } else {
        target = "xl/" + target;
      }
      rid2target->insert({rid, target});
    }
  }
  return 0;
}

static int ms_xlsx_fetch_sheet_bar_list(const std::string &xml,
                                        const xlsx_tags_t &t,
                                        std::vector<xlsx_sheet_bar_t> *sheets) {
  sheets->clear();

  const char *shts = nullptr;
  size_t shts_len = 0;
  const char *shts_end = nullptr;
  size_t shts_end_len = 0;
  if (!find_tag(xml.c_str(), t.sheets, &shts, &shts_len) ||
      !find_tag(shts + shts_len, t.sheets_, &shts_end, &shts_end_len)) {
    return -1;
  }
  _temp_truncate_string(shts, shts_end - shts);

  const char *sht = nullptr;
  size_t sht_len = 0;
  for (bool f = find_tag(shts + shts_len, t.sheet, &sht, &sht_len); f;
       f = find_tag(sht + sht_len, t.sheet, &sht, &sht_len)) {
    xlsx_sheet_bar_t s;
    if (get_attr_value(sht, sht_len, "r:id=", &s.rid) &&
        get_attr_value(sht, sht_len, "name=", &s.name)) {
      if (!get_attr_value(sht, sht_len, "state=", &s.state)) {
        s.state = "visible";
      }
      sheets->push_back(std::move(s));
    }
  }
  return 0;
}

int MsXLSxFetchSheetBarList(ZipHelper &zip, size_t xml_max_file_len,
                            std::vector<xlsx_sheet_bar_t> *sheets,
                            bool *is_xtag) {
  std::string xml;
  if (zip.ReadByName("xl/workbook.xml", xml_max_file_len, &xml) != 0) {
    return -1;
  }
  sheets->clear();

  if (ms_xlsx_fetch_sheet_bar_list(xml, g_xlsxTags, sheets) != 0) {
    if (is_xtag != nullptr) {
      *is_xtag = true;
    }
    return ms_xlsx_fetch_sheet_bar_list(xml, g_xlsxXTags, sheets);
  } else {
    if (is_xtag != nullptr) {
      *is_xtag = false;
    }
    return 0;
  }
}

static size_t ms_xlsx_fetch_sst(const std::string &xml, const xlsx_tags_t &tags,
                                int max_sst_cnt,
                                std::vector<std::string> *sst) {
  sst->clear();

  const char *si = nullptr;
  size_t si_len = 0;
  for (bool f = find_tag(xml.c_str(), tags.si, &si, &si_len);
       f && sst->size() < static_cast<size_t>(max_sst_cnt);
       f = find_tag(si + si_len, tags.si, &si, &si_len)) {
    const char *si_end = nullptr;
    size_t si_end_len = 0;
    if (!find_tag(si + si_len, tags.si_, &si_end, &si_end_len)) {
      break;
    }
    _temp_truncate_string(si, si_end - si);

    const char *t = nullptr;
    size_t t_len = 0;
    const char *t_end = nullptr;
    size_t t_end_len = 0;
    if (!find_tag(si + si_len, tags.t, &t, &t_len) ||
        !find_tag(t + t_len, tags.t_, &t_end, &t_end_len)) {
      break;
    }
    sst->push_back(std::string(t + t_len, t_end - (t + t_len)));
  }

  return sst->size();
}

int MsXLSxFetchSST(ZipHelper &zip, size_t xml_max_file_len, int max_sst_cnt,
                   std::vector<std::string> *sst) {
  std::string xml;
  if (zip.ReadByName("xl/sharedStrings.xml", xml_max_file_len, &xml) != 0) {
    return 0;
  }

  if (max_sst_cnt <= 0) {
    max_sst_cnt = std::numeric_limits<int>::max();
  }

  if (ms_xlsx_fetch_sst(xml, g_xlsxTags, max_sst_cnt, sst) == 0) {
    ms_xlsx_fetch_sst(xml, g_xlsxXTags, max_sst_cnt, sst);
  }
  return 0;
}

static int ms_xlsx_fetch_text(const char *xml,
                              const std::vector<std::string> &sst,
                              const std::string &delimiter,
                              const xlsx_tags_t &tags, size_t *max_len,
                              std::string *text) {
  size_t delimiter_len = utils::count_utf8_word_cnt(delimiter);
  const char *row = nullptr;
  size_t row_len = 0;
  for (bool f = find_tag(xml, tags.row, &row, &row_len); f && *max_len > 0;
       f = find_tag(row + row_len, tags.row, &row, &row_len)) {
    if (is_empty_element_tag(row, row_len)) {
      continue;
    }

    const char *row_end = nullptr;
    size_t row_end_size = 0;
    if (!find_tag(row + row_len, tags.row_, &row_end, &row_end_size)) {
      return -1;
    }
    _temp_truncate_string(row, row_end - row);

    bool empty_row = true;
    const char *c = nullptr;
    size_t c_len = 0;
    for (bool fc = find_tag(row + row_len, tags.c, &c, &c_len);
         fc && *max_len > 0 && c < row_end;
         fc = find_tag(c + c_len, tags.c, &c, &c_len)) {
      if (is_empty_element_tag(c, c_len)) {
        continue;
      }

      const char *c_end = nullptr;
      size_t c_end_len = 0;
      if (!find_tag(c + c_len, tags.c_, &c_end, &c_end_len)) {
        return -1;
      }
      _temp_truncate_string(c, c_end - c);

      const char *v = nullptr;
      size_t v_len = 0;
      const char *v_end = nullptr;
      size_t v_end_len = 0;
      if (!find_tag(c + c_len, tags.v, &v, &v_len) ||
          !find_tag(v + v_len, tags.v_, &v_end, &v_end_len)) {
        return -1;
      }
      std::string val(v + v_len, v_end - (v + v_len));

      static const std::string ignore_str = "_";
      const std::string *cell_text = &val;
      {
        _temp_truncate_string(c, c_len);
        if (is_sst_tag(c, c_len)) {
          size_t id = atoi(val.c_str());
          cell_text = id >= sst.size() ? &ignore_str : &sst[id];
        }
      }

      if (!empty_row) {
        append_text(text, max_len, delimiter.c_str(), delimiter.length(),
                    delimiter_len);
        if (max_len == 0) {
          return 0;
        }
      }
      append_text(text, max_len, cell_text->c_str(), cell_text->length());
      if (max_len == 0) {
        return 0;
      }
      empty_row = false;
    }

    if (*max_len > 0 && !empty_row) {
      text->push_back('\n');
      *max_len -= 1;
    }
  }
  return 0;
}

int MsXLSxFetchText(ZipHelper &zip, const fetch_text_options_t *opts,
                    std::string *text) {
  if (opts == nullptr) {
    opts = &__defaultFetchTextOptions;
  }

  std::vector<std::string> sst;
  if (MsXLSxFetchSST(zip, opts->xml_max_file_len, opts->xls_max_sst_cnt,
                     &sst) != 0) {
    return -1;
  }

  std::map<std::string, std::string> rid2target;
  if (MsXLSxFetchRelationships(zip, opts->xml_max_file_len, &rid2target) != 0) {
    return -1;
  }

  bool is_xtag;
  std::vector<xlsx_sheet_bar_t> sheets;
  if (MsXLSxFetchSheetBarList(zip, opts->xml_max_file_len, &sheets, &is_xtag) !=
      0) {
    return -1;
  }

  size_t max_len = opts->max_fetch_text_len;
  std::vector<char> name(128);
  std::string xml;
  for (auto &sht : sheets) {
    if (sht.state != "visible") {
      continue;
    }

    append_text(text, &max_len, sht.name.c_str(), sht.name.length());
    if (max_len > 0) {
      text->push_back('\n');
      max_len -= 1;
    }
    if (max_len == 0) {
      break;
    }

    auto it = rid2target.find(sht.rid);
    if (it == rid2target.end() || it->second.empty()) {
      continue;
    }
    if (zip.ReadByName(it->second, opts->xml_max_file_len, &xml) != 0) {
      continue;
    }

    ms_xlsx_fetch_text(xml.c_str(), sst, opts->xls_delimiter,
                       is_xtag ? g_xlsxXTags : g_xlsxTags, &max_len, text);

    if (max_len == 0) {
      break;
    }
  }

  return 0;
}

}  // namespace officex

}  // namespace msoffice
