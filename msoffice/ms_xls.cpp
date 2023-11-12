#include "msoffice/ms_xls.h"

#include <assert.h>
#include <stdio.h>
#include <string.h>

#include <codecvt>
#include <locale>
#include <unordered_map>

#include "utils/utils.h"

#define _STYLE_Info  "\e[3;32m"
#define _STYLE_Err   "\e[3;31m"
#define _STYLE_Warn  "\e[3;33m"
#define _STYLE_Debug "\e[3;36m"
#define slog(_type_, _fmt_, ...)                                      \
  printf(_STYLE_##_type_ "%.1s [%s:%s:%d]\e[0m " _fmt_ "\n", #_type_, \
         __FILE__, __func__, __LINE__, ##__VA_ARGS__)

namespace msoffice {

namespace xls {

static const std::string g_WorkbookName = "Workbook";
static const uint16_t g_BofVersion = 0x0600;
// static const size_t g_maxRecordSize = 8224;

const std::string &Identifier2Name(uint16_t identifier) {
  static const std::unordered_map<uint16_t, std::string> id2name = {
      {6, "Formula"},
      {10, "EOF"},
      {12, "CalcCount"},
      {13, "CalcMode"},
      {14, "CalcPrecision"},
      {15, "CalcRefMode"},
      {16, "CalcDelta"},
      {17, "CalcIter"},
      {18, "Protect"},
      {19, "Password"},
      {20, "Header"},
      {21, "Footer"},
      {23, "ExternSheet"},
      {24, "Lbl"},
      {25, "WinProtect"},
      {26, "VerticalPageBreaks"},
      {27, "HorizontalPageBreaks"},
      {28, "Note"},
      {29, "Selection"},
      {34, "Date1904"},
      {35, "ExternName"},
      {38, "LeftMargin"},
      {39, "RightMargin"},
      {40, "TopMargin"},
      {41, "BottomMargin"},
      {42, "PrintRowCol"},
      {43, "PrintGrid"},
      {47, "FilePass"},
      {49, "Font"},
      {51, "PrintSize"},
      {60, "Continue"},
      {61, "Window1"},
      {64, "Backup"},
      {65, "Pane"},
      {66, "CodePage"},
      {77, "Pls"},
      {80, "DCon"},
      {81, "DConRef"},
      {82, "DConName"},
      {85, "DefColWidth"},
      {89, "XCT"},
      {90, "CRN"},
      {91, "FileSharing"},
      {92, "WriteAccess"},
      {93, "Obj"},
      {94, "Uncalced"},
      {95, "CalcSaveRecalc"},
      {96, "Template"},
      {97, "Intl"},
      {99, "ObjProtect"},
      {125, "ColInfo"},
      {128, "Guts"},
      {129, "WsBool"},
      {130, "GridSet"},
      {131, "HCenter"},
      {132, "VCenter"},
      {133, "BoundSheet8"},
      {134, "WriteProtect"},
      {140, "Country"},
      {141, "HideObj"},
      {144, "Sort"},
      {146, "Palette"},
      {151, "Sync"},
      {152, "LPr"},
      {153, "DxGCol"},
      {154, "FnGroupName"},
      {155, "FilterMode"},
      {156, "BuiltInFnGroupCount"},
      {157, "AutoFilterInfo"},
      {158, "AutoFilter"},
      {160, "Scl"},
      {161, "Setup"},
      {174, "ScenMan"},
      {175, "SCENARIO"},
      {176, "SxView"},
      {177, "Sxvd"},
      {178, "SXVI"},
      {180, "SxIvd"},
      {181, "SXLI"},
      {182, "SXPI"},
      {184, "DocRoute"},
      {185, "RecipName"},
      {189, "MulRk"},
      {190, "MulBlank"},
      {193, "Mms"},
      {197, "SXDI"},
      {198, "SXDB"},
      {199, "SXFDB"},
      {200, "SXDBB"},
      {201, "SXNum"},
      {202, "SxBool"},
      {203, "SxErr"},
      {204, "SXInt"},
      {205, "SXString"},
      {206, "SXDtr"},
      {207, "SxNil"},
      {208, "SXTbl"},
      {209, "SXTBRGIITM"},
      {210, "SxTbpg"},
      {211, "ObProj"},
      {213, "SXStreamID"},
      {215, "DBCell"},
      {216, "SXRng"},
      {217, "SxIsxoper"},
      {218, "BookBool"},
      {220, "DbOrParamQry"},
      {221, "ScenarioProtect"},
      {222, "OleObjectSize"},
      {224, "XF"},
      {225, "InterfaceHdr"},
      {226, "InterfaceEnd"},
      {227, "SXVS"},
      {229, "MergeCells"},
      {233, "BkHim"},
      {235, "MsoDrawingGroup"},
      {236, "MsoDrawing"},
      {237, "MsoDrawingSelection"},
      {239, "PhoneticInfo"},
      {240, "SxRule"},
      {241, "SXEx"},
      {242, "SxFilt"},
      {244, "SxDXF"},
      {245, "SxItm"},
      {246, "SxName"},
      {247, "SxSelect"},
      {248, "SXPair"},
      {249, "SxFmla"},
      {251, "SxFormat"},
      {252, "SST"},
      {253, "LabelSst"},
      {255, "ExtSST"},
      {256, "SXVDEx"},
      {259, "SXFormula"},
      {290, "SXDBEx"},
      {311, "RRDInsDel"},
      {312, "RRDHead"},
      {315, "RRDChgCell"},
      {317, "RRTabId"},
      {318, "RRDRenSheet"},
      {319, "RRSort"},
      {320, "RRDMove"},
      {330, "RRFormat"},
      {331, "RRAutoFmt"},
      {333, "RRInsertSh"},
      {334, "RRDMoveBegin"},
      {335, "RRDMoveEnd"},
      {336, "RRDInsDelBegin"},
      {337, "RRDInsDelEnd"},
      {338, "RRDConflict"},
      {339, "RRDDefName"},
      {340, "RRDRstEtxp"},
      {351, "LRng"},
      {352, "UsesELFs"},
      {353, "DSF"},
      {401, "CUsr"},
      {402, "CbUsr"},
      {403, "UsrInfo"},
      {404, "UsrExcl"},
      {405, "FileLock"},
      {406, "RRDInfo"},
      {407, "BCUsrs"},
      {408, "UsrChk"},
      {425, "UserBView"},
      {426, "UserSViewBegin"},
      {427, "UserSViewEnd"},
      {428, "RRDUserView"},
      {429, "Qsi"},
      {430, "SupBook"},
      {431, "Prot4Rev"},
      {432, "CondFmt"},
      {433, "CF"},
      {434, "DVal"},
      {437, "DConBin"},
      {438, "TxO"},
      {439, "RefreshAll"},
      {440, "HLink"},
      {441, "Lel"},
      {442, "CodeName"},
      {443, "SXFDBType"},
      {444, "Prot4RevPass"},
      {445, "ObNoMacros"},
      {446, "Dv"},
      {448, "Excel9File"},
      {449, "RecalcId"},
      {450, "EntExU2"},
      {512, "Dimensions"},
      {513, "Blank"},
      {515, "Number"},
      {516, "Label"},
      {517, "BoolErr"},
      {519, "String"},
      {520, "Row"},
      {523, "Index"},
      {545, "Array"},
      {549, "DefaultRowHeight"},
      {566, "Table"},
      {574, "Window2"},
      {638, "RK"},
      {659, "Style"},
      {1048, "BigName"},
      {1054, "Format"},
      {1084, "ContinueBigName"},
      {1212, "ShrFmla"},
      {2048, "HLinkTooltip"},
      {2049, "WebPub"},
      {2050, "QsiSXTag"},
      {2051, "DBQueryExt"},
      {2052, "ExtString"},
      {2053, "TxtQry"},
      {2054, "Qsir"},
      {2055, "Qsif"},
      {2056, "RRDTQSIF"},
      {2057, "BOF"},
      {2058, "OleDbConn"},
      {2059, "WOpt"},
      {2060, "SXViewEx"},
      {2061, "SXTH"},
      {2062, "SXPIEx"},
      {2063, "SXVDTEx"},
      {2064, "SXViewEx9"},
      {2066, "ContinueFrt"},
      {2067, "RealTimeData"},
      {2128, "ChartFrtInfo"},
      {2129, "FrtWrapper"},
      {2130, "StartBlock"},
      {2131, "EndBlock"},
      {2132, "StartObject"},
      {2133, "EndObject"},
      {2134, "CatLab"},
      {2135, "YMult"},
      {2136, "SXViewLink"},
      {2137, "PivotChartBits"},
      {2138, "FrtFontList"},
      {2146, "SheetExt"},
      {2147, "BookExt"},
      {2148, "SXAddl"},
      {2149, "CrErr"},
      {2150, "HFPicture"},
      {2151, "FeatHdr"},
      {2152, "Feat"},
      {2154, "DataLabExt"},
      {2155, "DataLabExtContents"},
      {2156, "CellWatch"},
      {2161, "FeatHdr11"},
      {2162, "Feature11"},
      {2164, "DropDownObjIds"},
      {2165, "ContinueFrt11"},
      {2166, "DConn"},
      {2167, "List12"},
      {2168, "Feature12"},
      {2169, "CondFmt12"},
      {2170, "CF12"},
      {2171, "CFEx"},
      {2172, "XFCRC"},
      {2173, "XFExt"},
      {2174, "AutoFilter12"},
      {2175, "ContinueFrt12"},
      {2180, "MDTInfo"},
      {2181, "MDXStr"},
      {2182, "MDXTuple"},
      {2183, "MDXSet"},
      {2184, "MDXProp"},
      {2185, "MDXKPI"},
      {2186, "MDB"},
      {2187, "PLV"},
      {2188, "Compat12"},
      {2189, "DXF"},
      {2190, "TableStyles"},
      {2191, "TableStyle"},
      {2192, "TableStyleElement"},
      {2194, "StyleExt"},
      {2195, "NamePublish"},
      {2196, "NameCmt"},
      {2197, "SortData"},
      {2198, "Theme"},
      {2199, "GUIDTypeLib"},
      {2200, "FnGrp12"},
      {2201, "NameFnGrp12"},
      {2202, "MTRSettings"},
      {2203, "CompressPictures"},
      {2204, "HeaderFooter"},
      {2205, "CrtLayout12"},
      {2206, "CrtMlFrt"},
      {2207, "CrtMlFrtContinue"},
      {2211, "ForceFullCalculation"},
      {2212, "ShapePropsStream"},
      {2213, "TextPropsStream"},
      {2214, "RichTextStream"},
      {2215, "CrtLayout12A"},
      {4097, "Units"},
      {4098, "Chart"},
      {4099, "Series"},
      {4102, "DataFormat"},
      {4103, "LineFormat"},
      {4105, "MarkerFormat"},
      {4106, "AreaFormat"},
      {4107, "PieFormat"},
      {4108, "AttachedLabel"},
      {4109, "SeriesText"},
      {4116, "ChartFormat"},
      {4117, "Legend"},
      {4118, "SeriesList"},
      {4119, "Bar"},
      {4120, "Line"},
      {4121, "Pie"},
      {4122, "Area"},
      {4123, "Scatter"},
      {4124, "CrtLine"},
      {4125, "Axis"},
      {4126, "Tick"},
      {4127, "ValueRange"},
      {4128, "CatSerRange"},
      {4129, "AxisLine"},
      {4130, "CrtLink"},
      {4132, "DefaultText"},
      {4133, "Text"},
      {4134, "FontX"},
      {4135, "ObjectLink"},
      {4146, "Frame"},
      {4147, "Begin"},
      {4148, "End"},
      {4149, "PlotArea"},
      {4154, "Chart3d"},
      {4156, "PicF"},
      {4157, "DropBar"},
      {4158, "Radar"},
      {4159, "Surf"},
      {4160, "RadarArea"},
      {4161, "AxisParent"},
      {4163, "LegendException"},
      {4164, "ShtProps"},
      {4165, "SerToCrt"},
      {4166, "AxesUsed"},
      {4168, "SBaseRef"},
      {4170, "SerParent"},
      {4171, "SerAuxTrend"},
      {4174, "IFmtRecord"},
      {4175, "Pos"},
      {4176, "AlRuns"},
      {4177, "BRAI"},
      {4187, "SerAuxErrBar"},
      {4188, "ClrtClient"},
      {4189, "SerFmt"},
      {4191, "Chart3DBarShape"},
      {4192, "Fbi"},
      {4193, "BopPop"},
      {4194, "AxcExt"},
      {4195, "Dat"},
      {4196, "PlotGrowth"},
      {4197, "SIIndex"},
      {4198, "GelFrame"},
      {4199, "BopPopCustom"},
      {4200, "Fbi2"},
  };
  static const std::string unknown = "Unknown";
  auto it = id2name.find(identifier);
  return it == id2name.end() ? unknown : it->second;
}

// =============================================================================

template <typename T>
static inline int get_val_and_move(const char *data, size_t data_len,
                                   size_t *offset, T *val) {
  if (*offset + sizeof(T) > data_len) {
    return -1;
  }
  *val = *reinterpret_cast<const T *>(data + *offset);
  *offset += sizeof(T);
  return 0;
}

template <typename T>
static inline const T *get_ptr_and_move(const char *data, size_t data_len,
                                        size_t *offset) {
  if (*offset + sizeof(T) > data_len) {
    return nullptr;
  }
  auto ptr = reinterpret_cast<const T *>(data + *offset);
  *offset += sizeof(T);
  return ptr;
}

static int append_rgb_string(const char *data, size_t curr_block_end,
                             bool highbyte, size_t *offset, size_t *char_cnt,
                             std::string *rgb_str) {
  if (*offset > curr_block_end) {
    return -1;
  }

  size_t dcnt = 0;
  size_t rsize = curr_block_end - *offset;
  std::string tmp_str;
  if (highbyte) {  // char16
    size_t block_max_cnt = rsize / 2;
    dcnt = *char_cnt > block_max_cnt ? block_max_cnt : *char_cnt;
    auto begin = reinterpret_cast<const char16_t *>(data + *offset);
    if (Utf16ToUtf8(begin, begin + dcnt, &tmp_str) != 0) {
      return -1;
    }
    *offset += dcnt * 2;
    rgb_str->append(tmp_str);
  } else {
    dcnt = *char_cnt > rsize ? rsize : *char_cnt;
    tmp_str.assign(data + *offset, dcnt);
    *offset += dcnt;
    for (auto &ch : tmp_str) {
      if (ch < 0) {
        ch = ' ';
      }
    }
    rgb_str->append(tmp_str);
  }
  *char_cnt -= dcnt;
  return 0;
}

static int append_rgb_string_with_continue(const char *data, size_t data_len,
                                           size_t char_cnt, size_t *offset,
                                           size_t *curr_block_end,
                                           std::string *rgb_str) {
  for (; char_cnt > 0;) {
    if (*offset + sizeof(record_header_t) > data_len) {
      return -1;
    }

    auto next_rh = reinterpret_cast<const record_header_t *>(data + *offset);
    if (next_rh->identifier != kRecord_Continue ||
        *offset + next_rh->size > data_len) {
      return -1;
    }
    *curr_block_end = *offset + sizeof(record_header_t) + next_rh->size;
    *offset += sizeof(record_header_t);

    uint8_t flags;
    if (get_val_and_move(data, *curr_block_end, offset, &flags) != 0) {
      return -1;
    }

    if (append_rgb_string(data, *curr_block_end, flags & 0x1, offset, &char_cnt,
                          rgb_str) != 0) {
      return -1;
    }
  }
  return 0;
}

ssize_t FetchTextFromSST(const record_header_t &rh, const char *data,
                         size_t data_len, int max_sst_cnt,
                         std::vector<XLUnicodeRichExtendedString> *sst) {
  if (max_sst_cnt <= 0) {
    max_sst_cnt = std::numeric_limits<int>::max();
  }

  size_t curr_block_end = rh.size;
  if (curr_block_end > data_len) {
    return -1;
  }

  auto csTotal = reinterpret_cast<const int32_t *>(data);
  auto cstUnique = csTotal + 1;
  size_t offset = sizeof(int32_t) * 2;
  if (offset > curr_block_end || *csTotal < 0 || *cstUnique < 0) {
    slog(Err, "csTotal %d, cstUnique %d", *csTotal, *cstUnique);
    return -1;
  }
  // slog(Debug, "csTotal %d, cstUnique %d", *csTotal, *cstUnique);

  int sst_size = std::min(max_sst_cnt, *cstUnique);
  std::vector<char> tmp_buf;
  sst->reserve(sst_size);
  for (int i = 0; i < sst_size; ++i) {
    if (offset == curr_block_end) {
      auto next_rh = get_ptr_and_move<record_header_t>(data, data_len, &offset);
      if (next_rh == nullptr || next_rh->identifier != kRecord_Continue) {
        return -1;
      }
      curr_block_end += sizeof(record_header_t) + next_rh->size;
    }

    XLUnicodeRichExtendedString s;
    ssize_t ofs = s.ReadAndParse(data, data_len, offset, &curr_block_end);
    if (ofs < 0) {
      return -1;
    }
    offset += ofs;
    sst->push_back(std::move(s));
  }

  return max_sst_cnt >= *cstUnique ? offset : curr_block_end;
}

ssize_t ReadAndParse1stSubstream(
    const char *data, size_t data_len, int max_sst_cnt,
    std::vector<BoundSheet8> *bs,
    std::vector<XLUnicodeRichExtendedString> *sst) {
  size_t offset = 0;
  auto bof_rh = get_ptr_and_move<record_header_t>(data, data_len, &offset);
  if (bof_rh == nullptr || bof_rh->identifier != kRecord_BOF) {
    return -1;
  }

  auto bof = get_ptr_and_move<BOF_t>(data, data_len, &offset);
  if (bof == nullptr || bof->vers != g_BofVersion) {
    return -1;
  }

  for (; offset < data_len;) {
    auto rh = get_ptr_and_move<record_header_t>(data, data_len, &offset);
    if (rh == nullptr) {
      return -1;
    }
    // slog(Info, "%s, size %d", Identifier2Name(rh->identifier).c_str(),
    //       rh->size);

    if (rh->identifier == kRecord_EOF) {
      return offset;
    } else if (rh->identifier == kRecord_SST) {
      ssize_t ofs = FetchTextFromSST(*rh, data + offset, data_len - offset,
                                     max_sst_cnt, sst);
      if (ofs < 0) {
        return -1;
      }
      offset += ofs;
      continue;
    } else if (rh->identifier == kRecord_BoundSheet8) {
      BoundSheet8 b;
      if (b.ParseFrom(data + offset, data_len - offset) != 0) {
        return -1;
      }
      bs->push_back(std::move(b));
    } else if (rh->identifier == kRecord_FilePass) {
      slog(Err, "file was encrypted");
      return -1;
    }

    offset += rh->size;
  }

  return -1;
}

// =============================================================================

ssize_t XLUnicodeRichExtendedString::ReadAndParse(const char *data,
                                                  size_t data_len,
                                                  size_t offset,
                                                  size_t *curr_block_end) {
  size_t ofs = offset;
  auto hdr = get_ptr_and_move<hdr_t>(data, *curr_block_end, &ofs);
  if (hdr == nullptr) {
    return -1;
  }
  m_hdr = *hdr;

  uint16_t cRun = 0;
  int32_t cbExtRst = 0;
  if (m_hdr.fRichSt() == 0x1) {
    if (get_val_and_move(data, *curr_block_end, &ofs, &cRun) != 0) {
      return -1;
    }
  }
  if (m_hdr.fExtSt() == 0x1) {
    if (get_val_and_move(data, *curr_block_end, &ofs, &cbExtRst) != 0) {
      return -1;
    }
  }
  // slog(Debug, "offset %ld, cRun %d, cbExtRst %d, hb %d", ofs, cRun,
  //       cbExtRst, m_hdr.fHighByte());

  // rgb
  size_t char_cnt = m_hdr.cch;
  if (append_rgb_string(data, *curr_block_end, m_hdr.fHighByte(), &ofs,
                        &char_cnt, &m_str) != 0) {
    return -1;
  }
  if (char_cnt > 0) {
    if (append_rgb_string_with_continue(data, data_len, char_cnt, &ofs,
                                        curr_block_end, &m_str) != 0) {
      return -1;
    }
  }

  // rgRun (optional)
  if (m_hdr.fRichSt() == 0x1) {
    ofs += 4 * cRun;
  }

  // ExtRst (optional)
  if (m_hdr.fExtSt() == 0x1) {
    ofs += cbExtRst;
  }

  return ofs - offset;
}

double RkNumber_t::value() const {
  if (fInt() == 1) {
    uint32_t n = num();
    auto val = reinterpret_cast<int *>(&n);
    return fX100() ? *val / 100.0 : *val;

  } else {
    uint64_t n = num();
    n <<= 34;
    auto val = reinterpret_cast<double *>(&n);
    return fX100() ? *val / 100 : *val;
  }
}

int MulRk::ParseFrom(const char *data, size_t data_len) {
  size_t offset = 0;

  if (get_val_and_move(data, data_len, &offset, &m_rw) != 0) {
    return -1;
  }

  if (get_val_and_move(data, data_len, &offset, &m_colFirst) != 0 ||
      m_colFirst > 254) {
    return -1;
  }

  if (data_len - offset < sizeof(uint16_t)) {
    return -1;
  }

  m_colLast =
      *reinterpret_cast<const uint16_t *>(data + (data_len - sizeof(uint16_t)));

  size_t size = data_len - offset - sizeof(uint16_t);
  if (size % sizeof(RkRec_t) != 0 ||
      size / sizeof(RkRec_t) != m_colLast - m_colFirst + 1) {
    return -1;
  }

  m_rgrkrec.resize(size / sizeof(RkRec_t));
  memcpy(m_rgrkrec.data(), data + offset, size);

  return 0;
}

int MulBlank::ParseFrom(const char *data, size_t data_len) {
  size_t offset = 0;

  if (get_val_and_move(data, data_len, &offset, &m_rw) != 0) {
    return -1;
  }
  if (get_val_and_move(data, data_len, &offset, &m_colFirst) != 0 ||
      m_colFirst > 254) {
    return -1;
  }

  if (data_len - offset < sizeof(uint16_t)) {
    return -1;
  }

  m_colLast =
      *reinterpret_cast<const uint16_t *>(data + (data_len - sizeof(uint16_t)));

  size_t size = data_len - offset - sizeof(uint16_t);
  if (size % sizeof(uint16_t) != 0 ||
      size / sizeof(uint16_t) != m_colLast - m_colFirst + 1) {
    return -1;
  }
  return 0;
}

ssize_t ShortXLUnicodeString::ReadAndParse(const char *data, size_t data_len) {
  size_t offset = 0;

  if (get_val_and_move(data, data_len, &offset, &m_cch) != 0) {
    return -1;
  }

  uint8_t fHighByte;
  if (get_val_and_move(data, data_len, &offset, &fHighByte) != 0) {
    return -1;
  }
  fHighByte = fHighByte & 0x1;

  if (fHighByte) {
    if (offset + m_cch * 2 > data_len) {
      return -1;
    }
    auto begin = reinterpret_cast<const char16_t *>(data + offset);
    auto end = begin + m_cch;
    if (Utf16ToUtf8(begin, end, &m_str) != 0) {
      return -1;
    }
    offset += m_cch * 2;
  } else {
    if (offset + m_cch > data_len) {
      return -1;
    }
    m_str.assign(data + offset, m_cch);
    offset += m_cch;
  }

  return offset;
}

int BoundSheet8::ParseFrom(const char *data, size_t data_len) {
  size_t offset = 0;

  if (get_val_and_move(data, data_len, &offset, &m_lbPlyPos) != 0) {
    return -1;
  }

  uint8_t hsState;
  if (get_val_and_move(data, data_len, &offset, &hsState) != 0) {
    return -1;
  }
  m_hsState = hsState & 0b11;

  if (get_val_and_move(data, data_len, &offset, &m_dt) != 0) {
    return -1;
  }

  ssize_t ofs = m_name.ReadAndParse(data + offset, data_len - offset);
  if (ofs < 0 || offset + ofs > data_len) {
    return -1;
  }

  return 0;
}

// =============================================================================

int MsXLS::ParseFromFile(const std::string &filename) {
  std::vector<char> data;
  utils::read_file(filename.c_str(), &data);
  if (data.size() < sizeof(compound_doc_header_t)) {
    return -1;
  }
  return ParseFromBytes(std::move<>(data));
}

int MsXLS::parse() {
  m_idx_workbook = -1;
  auto &dirs = m_comp_doc.GetDirEntries();
  for (int i = 0; i < dirs.size(); ++i) {
    if (dirs[i].type == kDirEntryTypeEmpty) {
      continue;
    }

    std::string dirname;
    if (Utf16ToUtf8(dirs[i].unicode_name,
                    dirs[i].unicode_name + dirs[i].name_len / 2 - 1,
                    &dirname) != 0) {
      continue;
    }

    if (dirname == g_WorkbookName) {
      m_idx_workbook = i;
    }
  }
  if (m_idx_workbook == -1) {
    return -1;
  }

  // if (m_comp_doc.GetDirEntryStream(dirs[m_idx_workbook], &m_workbook_stream)
  // !=
  //     0) {
  //   return -1;
  // }
  // if (ReadAndParse1stSubstream(m_workbook_stream.data(),
  //                              m_workbook_stream.size(), &m_bs_list,
  //                              &m_sst) < 0) {
  //   return -1;
  // }

  return 0;
}

static int append_cell(const char *str, size_t str_cch, const char *delimiter,
                       size_t delimiter_cch, uint16_t cell_row,
                       uint16_t cell_col, uint16_t *curr_row,
                       size_t *max_fetch_text_len, std::string *text) {
  if (*max_fetch_text_len == 0) {
    return 1;
  }

  if (cell_row != *curr_row) {
    text->push_back('\n');
    *curr_row = cell_row;
    *max_fetch_text_len -= 1;

    if (*max_fetch_text_len == 0) {
      return 1;
    }
  }

  if (cell_col != 0) {
    if (*max_fetch_text_len < delimiter_cch) {
      text->append(delimiter,
                   utils::fix_utf8_word_cnt(delimiter, *max_fetch_text_len));
      *max_fetch_text_len = 0;
      return 1;
    } else {
      text->append(delimiter);
      *max_fetch_text_len -= delimiter_cch;
    }

    if (*max_fetch_text_len == 0) {
      return 1;
    }
  }

  if (*max_fetch_text_len < str_cch) {
    text->append(str, utils::fix_utf8_word_cnt(str, *max_fetch_text_len));
    *max_fetch_text_len = 0;
    return 1;
  } else {
    text->append(str);
    *max_fetch_text_len -= str_cch;
  }

  return *max_fetch_text_len == 0 ? 1 : 0;
}

static void to_string(double f, std::vector<char> *buf) {
  if (f > 10) {
    snprintf(buf->data(), buf->size(), "%.2f", f);
  } else {
    snprintf(buf->data(), buf->size(), "%f", f);
  }
}

static void to_string(const RkNumber_t &rk, std::vector<char> *buf) {
  snprintf(buf->data(), buf->size(), "%.2f", rk.value());
}

int MsXLS::FetchText(const fetch_text_options_t *user_opts,
                     std::string *text) const {
  static const fetch_text_options_t default_opts;
  fetch_text_options_t opts = user_opts != nullptr ? *user_opts : default_opts;
  size_t delimiter_cch = utils::count_utf8_word_cnt(opts.xls_delimiter);

  std::vector<char> workbook_stream;
  if (m_comp_doc.GetDirEntryStream(m_comp_doc.GetDirEntries()[m_idx_workbook],
                                   &workbook_stream) != 0) {
    return -1;
  }
  const char *data = workbook_stream.data();
  size_t data_len = workbook_stream.size();

  std::vector<BoundSheet8> bs_list;
  std::vector<XLUnicodeRichExtendedString> sst;
  if (ReadAndParse1stSubstream(data, data_len, opts.xls_max_sst_cnt, &bs_list,
                               &sst) < 0) {
    return -1;
  }

  // const char *data = m_workbook_stream.data();
  // size_t data_len = m_workbook_stream.size();
  // auto &bs_list = m_bs_list;
  // auto &sst = m_sst;

  for (auto &bs : bs_list) {
    if (bs.Dt() != BoundSheet8::kDT_WorksheetOrDialogSheet ||
        bs.HsState() != 0x00) {
      continue;
    }

    if (opts.max_fetch_text_len < bs.Name().Cch()) {
      auto &name = bs.Name().String();
      text->append(name.c_str(), utils::fix_utf8_word_cnt(
                                     name.c_str(), opts.max_fetch_text_len));
      opts.max_fetch_text_len = 0;
    } else {
      text->append(bs.Name().String().c_str());
      opts.max_fetch_text_len -= bs.Name().Cch();
    }
    if (opts.max_fetch_text_len == 0) {
      break;
    }

    text->push_back('\n');
    opts.max_fetch_text_len -= 1;
    if (opts.max_fetch_text_len == 0) {
      break;
    }

    size_t offset = bs.LbPlyPos();
    if (offset > data_len || offset + sizeof(record_header_t) > data_len) {
      return -1;
    }

#define _get_ptr(_ptr, _type)                                   \
  auto _ptr = get_ptr_and_move<_type>(data, data_len, &offset); \
  if (_ptr == nullptr) {                                        \
    return -1;                                                  \
  }

    _get_ptr(bof_rh, record_header_t);
    if (bof_rh->identifier != kRecord_BOF) {
      return -1;
    }

    _get_ptr(bof, BOF_t);
    bool eof = false;
    uint16_t row = 0;
    std::vector<char> buf(64);
    for (; offset < data_len && opts.max_fetch_text_len > 0;) {
      _get_ptr(rh, record_header_t);

      if (rh->identifier == kRecord_EOF) {
        text->push_back('\n');
        opts.max_fetch_text_len -= 1;
        eof = true;
        break;
      } else if (rh->identifier == kRecord_LabelSst) {
        _get_ptr(lab, LabelSst_t);
        const char *sst_txt = "_";
        uint16_t sst_txt_cch = 1;
        if (lab->isst < sst.size()) {
          sst_txt = sst[lab->isst].String().c_str();
          sst_txt_cch = sst[lab->isst].Hdr().cch;
        }
        append_cell(sst_txt, sst_txt_cch, opts.xls_delimiter.c_str(),
                    delimiter_cch, lab->cell.rw, lab->cell.col, &row,
                    &opts.max_fetch_text_len, text);

      } else if (rh->identifier == kRecord_RK) {
        _get_ptr(rk, RK_t);
        to_string(rk->rkrec.RK, &buf);
        append_cell(buf.data(), strlen(buf.data()), opts.xls_delimiter.c_str(),
                    delimiter_cch, rk->rw, rk->col, &row,
                    &opts.max_fetch_text_len, text);

      } else if (rh->identifier == kRecord_MulRk) {
        MulRk mrk;
        if (mrk.ParseFrom(data + offset, rh->size) != 0) {
          return -1;
        }
        uint16_t cell_col = mrk.ColFirst();
        for (auto &rk : mrk.RgRkrec()) {
          to_string(rk.RK, &buf);
          append_cell(buf.data(), strlen(buf.data()),
                      opts.xls_delimiter.c_str(), delimiter_cch, mrk.Rw(),
                      cell_col++, &row, &opts.max_fetch_text_len, text);
        }
        offset += rh->size;
      } else if (rh->identifier == kRecord_Number) {
        auto n = reinterpret_cast<const Number_t *>(data + offset);
        to_string(n->num, &buf);
        append_cell(buf.data(), strlen(buf.data()), opts.xls_delimiter.c_str(),
                    delimiter_cch, n->cell.rw, n->cell.col, &row,
                    &opts.max_fetch_text_len, text);
        offset += rh->size;
      } else if (rh->identifier == kRecord_Blank && !opts.xls_skip_blank_cell) {
        _get_ptr(bk, Blank_t);
        append_cell(" ", 1, opts.xls_delimiter.c_str(), delimiter_cch,
                    bk->cell.rw, bk->cell.col, &row, &opts.max_fetch_text_len,
                    text);
      } else if (rh->identifier == kRecord_MulBlank &&
                 !opts.xls_skip_blank_cell) {
        MulBlank mbk;
        if (mbk.ParseFrom(data + offset, rh->size) != 0) {
          return -1;
        }
        for (uint16_t cell_col = mbk.ColFirst(); cell_col < mbk.ColLast();
             ++cell_col) {
          append_cell(" ", 1, opts.xls_delimiter.c_str(), delimiter_cch,
                      mbk.Rw(), cell_col, &row, &opts.max_fetch_text_len, text);
        }
        offset += rh->size;
      } else {
        offset += rh->size;
      }
    }
    if (!eof && opts.max_fetch_text_len != 0) {
      return -1;
    }
  }
  RemoveControlCharacter(text);
  return 0;

#undef _get_ptr
}

}  // namespace xls

}  // namespace msoffice
