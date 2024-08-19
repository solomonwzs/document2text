#pragma once

#include <string>
#include <vector>

#include "msoffice/compound_document.h"
#include "msoffice/ms_xls/crypt/Crypt.h"
#include "msoffice/utils.h"

namespace msoffice {

namespace xls {

enum record_t {
  kRecord_Formula = 6,
  kRecord_EOF = 10,
  kRecord_CalcCount = 12,
  kRecord_CalcMode = 13,
  kRecord_CalcPrecision = 14,
  kRecord_CalcRefMode = 15,
  kRecord_CalcDelta = 16,
  kRecord_CalcIter = 17,
  kRecord_Protect = 18,
  kRecord_Password = 19,
  kRecord_Header = 20,
  kRecord_Footer = 21,
  kRecord_ExternSheet = 23,
  kRecord_Lbl = 24,
  kRecord_WinProtect = 25,
  kRecord_VerticalPageBreaks = 26,
  kRecord_HorizontalPageBreaks = 27,
  kRecord_Note = 28,
  kRecord_Selection = 29,
  kRecord_Date1904 = 34,
  kRecord_ExternName = 35,
  kRecord_LeftMargin = 38,
  kRecord_RightMargin = 39,
  kRecord_TopMargin = 40,
  kRecord_BottomMargin = 41,
  kRecord_PrintRowCol = 42,
  kRecord_PrintGrid = 43,
  kRecord_FilePass = 47,
  kRecord_Font = 49,
  kRecord_PrintSize = 51,
  kRecord_Continue = 60,
  kRecord_Window1 = 61,
  kRecord_Backup = 64,
  kRecord_Pane = 65,
  kRecord_CodePage = 66,
  kRecord_Pls = 77,
  kRecord_DCon = 80,
  kRecord_DConRef = 81,
  kRecord_DConName = 82,
  kRecord_DefColWidth = 85,
  kRecord_XCT = 89,
  kRecord_CRN = 90,
  kRecord_FileSharing = 91,
  kRecord_WriteAccess = 92,
  kRecord_Obj = 93,
  kRecord_Uncalced = 94,
  kRecord_CalcSaveRecalc = 95,
  kRecord_Template = 96,
  kRecord_Intl = 97,
  kRecord_ObjProtect = 99,
  kRecord_ColInfo = 125,
  kRecord_Guts = 128,
  kRecord_WsBool = 129,
  kRecord_GridSet = 130,
  kRecord_HCenter = 131,
  kRecord_VCenter = 132,
  kRecord_BoundSheet8 = 133,
  kRecord_WriteProtect = 134,
  kRecord_Country = 140,
  kRecord_HideObj = 141,
  kRecord_Sort = 144,
  kRecord_Palette = 146,
  kRecord_Sync = 151,
  kRecord_LPr = 152,
  kRecord_DxGCol = 153,
  kRecord_FnGroupName = 154,
  kRecord_FilterMode = 155,
  kRecord_BuiltInFnGroupCount = 156,
  kRecord_AutoFilterInfo = 157,
  kRecord_AutoFilter = 158,
  kRecord_Scl = 160,
  kRecord_Setup = 161,
  kRecord_ScenMan = 174,
  kRecord_SCENARIO = 175,
  kRecord_SxView = 176,
  kRecord_Sxvd = 177,
  kRecord_SXVI = 178,
  kRecord_SxIvd = 180,
  kRecord_SXLI = 181,
  kRecord_SXPI = 182,
  kRecord_DocRoute = 184,
  kRecord_RecipName = 185,
  kRecord_MulRk = 189,
  kRecord_MulBlank = 190,
  kRecord_Mms = 193,
  kRecord_SXDI = 197,
  kRecord_SXDB = 198,
  kRecord_SXFDB = 199,
  kRecord_SXDBB = 200,
  kRecord_SXNum = 201,
  kRecord_SxBool = 202,
  kRecord_SxErr = 203,
  kRecord_SXInt = 204,
  kRecord_SXString = 205,
  kRecord_SXDtr = 206,
  kRecord_SxNil = 207,
  kRecord_SXTbl = 208,
  kRecord_SXTBRGIITM = 209,
  kRecord_SxTbpg = 210,
  kRecord_ObProj = 211,
  kRecord_SXStreamID = 213,
  kRecord_DBCell = 215,
  kRecord_SXRng = 216,
  kRecord_SxIsxoper = 217,
  kRecord_BookBool = 218,
  kRecord_DbOrParamQry = 220,
  kRecord_ScenarioProtect = 221,
  kRecord_OleObjectSize = 222,
  kRecord_XF = 224,
  kRecord_InterfaceHdr = 225,
  kRecord_InterfaceEnd = 226,
  kRecord_SXVS = 227,
  kRecord_MergeCells = 229,
  kRecord_BkHim = 233,
  kRecord_MsoDrawingGroup = 235,
  kRecord_MsoDrawing = 236,
  kRecord_MsoDrawingSelection = 237,
  kRecord_PhoneticInfo = 239,
  kRecord_SxRule = 240,
  kRecord_SXEx = 241,
  kRecord_SxFilt = 242,
  kRecord_SxDXF = 244,
  kRecord_SxItm = 245,
  kRecord_SxName = 246,
  kRecord_SxSelect = 247,
  kRecord_SXPair = 248,
  kRecord_SxFmla = 249,
  kRecord_SxFormat = 251,
  kRecord_SST = 252,
  kRecord_LabelSst = 253,
  kRecord_ExtSST = 255,
  kRecord_SXVDEx = 256,
  kRecord_SXFormula = 259,
  kRecord_SXDBEx = 290,
  kRecord_RRDInsDel = 311,
  kRecord_RRDHead = 312,
  kRecord_RRDChgCell = 315,
  kRecord_RRTabId = 317,
  kRecord_RRDRenSheet = 318,
  kRecord_RRSort = 319,
  kRecord_RRDMove = 320,
  kRecord_RRFormat = 330,
  kRecord_RRAutoFmt = 331,
  kRecord_RRInsertSh = 333,
  kRecord_RRDMoveBegin = 334,
  kRecord_RRDMoveEnd = 335,
  kRecord_RRDInsDelBegin = 336,
  kRecord_RRDInsDelEnd = 337,
  kRecord_RRDConflict = 338,
  kRecord_RRDDefName = 339,
  kRecord_RRDRstEtxp = 340,
  kRecord_LRng = 351,
  kRecord_UsesELFs = 352,
  kRecord_DSF = 353,
  kRecord_CUsr = 401,
  kRecord_CbUsr = 402,
  kRecord_UsrInfo = 403,
  kRecord_UsrExcl = 404,
  kRecord_FileLock = 405,
  kRecord_RRDInfo = 406,
  kRecord_BCUsrs = 407,
  kRecord_UsrChk = 408,
  kRecord_UserBView = 425,
  kRecord_UserSViewBegin = 426,
  kRecord_UserSViewBegin_Chart = 426,
  kRecord_UserSViewEnd = 427,
  kRecord_RRDUserView = 428,
  kRecord_Qsi = 429,
  kRecord_SupBook = 430,
  kRecord_Prot4Rev = 431,
  kRecord_CondFmt = 432,
  kRecord_CF = 433,
  kRecord_DVal = 434,
  kRecord_DConBin = 437,
  kRecord_TxO = 438,
  kRecord_RefreshAll = 439,
  kRecord_HLink = 440,
  kRecord_Lel = 441,
  kRecord_CodeName = 442,
  kRecord_SXFDBType = 443,
  kRecord_Prot4RevPass = 444,
  kRecord_ObNoMacros = 445,
  kRecord_Dv = 446,
  kRecord_Excel9File = 448,
  kRecord_RecalcId = 449,
  kRecord_EntExU2 = 450,
  kRecord_Dimensions = 512,
  kRecord_Blank = 513,
  kRecord_Number = 515,
  kRecord_Label = 516,
  kRecord_BoolErr = 517,
  kRecord_String = 519,
  kRecord_Row = 520,
  kRecord_Index = 523,
  kRecord_Array = 545,
  kRecord_DefaultRowHeight = 549,
  kRecord_Table = 566,
  kRecord_Window2 = 574,
  kRecord_RK = 638,
  kRecord_Style = 659,
  kRecord_BigName = 1048,
  kRecord_Format = 1054,
  kRecord_ContinueBigName = 1084,
  kRecord_ShrFmla = 1212,
  kRecord_HLinkTooltip = 2048,
  kRecord_WebPub = 2049,
  kRecord_QsiSXTag = 2050,
  kRecord_DBQueryExt = 2051,
  kRecord_ExtString = 2052,
  kRecord_TxtQry = 2053,
  kRecord_Qsir = 2054,
  kRecord_Qsif = 2055,
  kRecord_RRDTQSIF = 2056,
  kRecord_BOF = 2057,
  kRecord_OleDbConn = 2058,
  kRecord_WOpt = 2059,
  kRecord_SXViewEx = 2060,
  kRecord_SXTH = 2061,
  kRecord_SXPIEx = 2062,
  kRecord_SXVDTEx = 2063,
  kRecord_SXViewEx9 = 2064,
  kRecord_ContinueFrt = 2066,
  kRecord_RealTimeData = 2067,
  kRecord_ChartFrtInfo = 2128,
  kRecord_FrtWrapper = 2129,
  kRecord_StartBlock = 2130,
  kRecord_EndBlock = 2131,
  kRecord_StartObject = 2132,
  kRecord_EndObject = 2133,
  kRecord_CatLab = 2134,
  kRecord_YMult = 2135,
  kRecord_SXViewLink = 2136,
  kRecord_PivotChartBits = 2137,
  kRecord_FrtFontList = 2138,
  kRecord_SheetExt = 2146,
  kRecord_BookExt = 2147,
  kRecord_SXAddl = 2148,
  kRecord_CrErr = 2149,
  kRecord_HFPicture = 2150,
  kRecord_FeatHdr = 2151,
  kRecord_Feat = 2152,
  kRecord_DataLabExt = 2154,
  kRecord_DataLabExtContents = 2155,
  kRecord_CellWatch = 2156,
  kRecord_FeatHdr11 = 2161,
  kRecord_Feature11 = 2162,
  kRecord_DropDownObjIds = 2164,
  kRecord_ContinueFrt11 = 2165,
  kRecord_DConn = 2166,
  kRecord_List12 = 2167,
  kRecord_Feature12 = 2168,
  kRecord_CondFmt12 = 2169,
  kRecord_CF12 = 2170,
  kRecord_CFEx = 2171,
  kRecord_XFCRC = 2172,
  kRecord_XFExt = 2173,
  kRecord_AutoFilter12 = 2174,
  kRecord_ContinueFrt12 = 2175,
  kRecord_MDTInfo = 2180,
  kRecord_MDXStr = 2181,
  kRecord_MDXTuple = 2182,
  kRecord_MDXSet = 2183,
  kRecord_MDXProp = 2184,
  kRecord_MDXKPI = 2185,
  kRecord_MDB = 2186,
  kRecord_PLV = 2187,
  kRecord_Compat12 = 2188,
  kRecord_DXF = 2189,
  kRecord_TableStyles = 2190,
  kRecord_TableStyle = 2191,
  kRecord_TableStyleElement = 2192,
  kRecord_StyleExt = 2194,
  kRecord_NamePublish = 2195,
  kRecord_NameCmt = 2196,
  kRecord_SortData = 2197,
  kRecord_Theme = 2198,
  kRecord_GUIDTypeLib = 2199,
  kRecord_FnGrp12 = 2200,
  kRecord_NameFnGrp12 = 2201,
  kRecord_MTRSettings = 2202,
  kRecord_CompressPictures = 2203,
  kRecord_HeaderFooter = 2204,
  kRecord_CrtLayout12 = 2205,
  kRecord_CrtMlFrt = 2206,
  kRecord_CrtMlFrtContinue = 2207,
  kRecord_ForceFullCalculation = 2211,
  kRecord_ShapePropsStream = 2212,
  kRecord_TextPropsStream = 2213,
  kRecord_RichTextStream = 2214,
  kRecord_CrtLayout12A = 2215,
  kRecord_Units = 4097,
  kRecord_Chart = 4098,
  kRecord_Series = 4099,
  kRecord_DataFormat = 4102,
  kRecord_LineFormat = 4103,
  kRecord_MarkerFormat = 4105,
  kRecord_AreaFormat = 4106,
  kRecord_PieFormat = 4107,
  kRecord_AttachedLabel = 4108,
  kRecord_SeriesText = 4109,
  kRecord_ChartFormat = 4116,
  kRecord_Legend = 4117,
  kRecord_SeriesList = 4118,
  kRecord_Bar = 4119,
  kRecord_Line = 4120,
  kRecord_Pie = 4121,
  kRecord_Area = 4122,
  kRecord_Scatter = 4123,
  kRecord_CrtLine = 4124,
  kRecord_Axis = 4125,
  kRecord_Tick = 4126,
  kRecord_ValueRange = 4127,
  kRecord_CatSerRange = 4128,
  kRecord_AxisLine = 4129,
  kRecord_CrtLink = 4130,
  kRecord_DefaultText = 4132,
  kRecord_Text = 4133,
  kRecord_FontX = 4134,
  kRecord_ObjectLink = 4135,
  kRecord_Frame = 4146,
  kRecord_Begin = 4147,
  kRecord_End = 4148,
  kRecord_PlotArea = 4149,
  kRecord_Chart3d = 4154,
  kRecord_PicF = 4156,
  kRecord_DropBar = 4157,
  kRecord_Radar = 4158,
  kRecord_Surf = 4159,
  kRecord_RadarArea = 4160,
  kRecord_AxisParent = 4161,
  kRecord_LegendException = 4163,
  kRecord_ShtProps = 4164,
  kRecord_SerToCrt = 4165,
  kRecord_AxesUsed = 4166,
  kRecord_SBaseRef = 4168,
  kRecord_SerParent = 4170,
  kRecord_SerAuxTrend = 4171,
  kRecord_IFmtRecord = 4174,
  kRecord_Pos = 4175,
  kRecord_AlRuns = 4176,
  kRecord_BRAI = 4177,
  kRecord_SerAuxErrBar = 4187,
  kRecord_ClrtClient = 4188,
  kRecord_SerFmt = 4189,
  kRecord_Chart3DBarShape = 4191,
  kRecord_Fbi = 4192,
  kRecord_BopPop = 4193,
  kRecord_AxcExt = 4194,
  kRecord_Dat = 4195,
  kRecord_PlotGrowth = 4196,
  kRecord_SIIndex = 4197,
  kRecord_GelFrame = 4198,
  kRecord_BopPopCustom = 4199,
  kRecord_Fbi2 = 4200,
};

struct record_header_t {
  uint16_t identifier;
  uint16_t size;
} __attribute__((packed));

enum document_type_t {
  kDT_WorkbookStream = 0x0005,
  kDT_DialogOrWorksheetSubstream = 0x0010,
  kDT_ChatSheetSubstream = 0x0020,
  kDT_MacroSheetSubstream = 0x0040,
};

struct rc4_encryption_header_t {
  int16_t vMajor;
  int16_t vMinor;
  CRYPT::_rc4CryptData data;
} __attribute__((packed));

struct BOF_t {
  uint16_t vers;
  uint16_t bt;
  uint16_t rupBuild;
  uint16_t rupYear;
  uint32_t flags1;
  uint8_t verLowestBiff;
  uint8_t flags2;
  uint16_t reserved;

#define _nth_bit(_f, _nth) (((_f) & (1u << (_nth))) >> _nth)
  inline uint8_t fWin() const {
    return _nth_bit(flags1, 0);
  }
  inline uint8_t fRisc() const {
    return _nth_bit(flags1, 1);
  }
  inline uint8_t fBeta() const {
    return _nth_bit(flags1, 2);
  }
  inline uint8_t fWinAny() const {
    return _nth_bit(flags1, 3);
  }
  inline uint8_t fMacAny() const {
    return _nth_bit(flags1, 4);
  }
  inline uint8_t fBetaAny() const {
    return _nth_bit(flags1, 5);
  }
  inline uint8_t fRiscAny() const {
    return _nth_bit(flags1, 8);
  }
  inline uint8_t fOOM() const {
    return _nth_bit(flags1, 9);
  }
  inline uint8_t fGlJmp() const {
    return _nth_bit(flags1, 10);
  }
  inline uint8_t fFontLimit() const {
    return _nth_bit(flags1, 13);
  }
  inline uint8_t verXLHigh() const {
    return (flags1 >> 14) & 0b1111;
  }
  inline uint8_t verLastXLSaved() const {
    return flags2 & 0b1111;
  }
#undef _nth_bit
} __attribute__((packed));

struct Row_t {
  uint16_t rw;
  uint16_t colMic;
  uint16_t colMac;
  uint16_t miyRw;
  uint16_t reserved1;
  uint16_t unused1;
  uint8_t flags1;
  uint8_t reserved3;
  uint16_t flags2;

#define _nth_bit(_f, _nth) (((_f) & (1u << (_nth))) >> _nth)
  inline uint8_t iOutLevel() const {
    return flags1 & 0b111;
  }
  inline uint8_t fCollapsed() const {
    return _nth_bit(flags1, 4);
  }
  inline uint8_t fDyZero() const {
    return _nth_bit(flags1, 5);
  }
  inline uint8_t fUnsynced() const {
    return _nth_bit(flags1, 6);
  }
  inline uint8_t fGhostDirty() const {
    return _nth_bit(flags1, 7);
  }

  inline uint16_t ixfe_val() const {
    return flags2 & 0b111111111111;
  }
  inline uint8_t fExAsc() const {
    return _nth_bit(flags2, 12);
  }
  inline uint8_t fExDes() const {
    return _nth_bit(flags2, 13);
  }
  inline uint8_t fPhonetic() const {
    return _nth_bit(flags2, 14);
  }
#undef _nth_bit
} __attribute__((packed));

struct Dimensions_t {
  uint32_t rwMic;
  uint32_t rwMac;
  uint16_t colMic;
  uint16_t colMac;
  uint16_t reserved;
} __attribute__((packed));

struct Cell_t {
  uint16_t rw;
  uint16_t col;
  uint16_t ixfe;
} __attribute__((packed));

struct Blank_t {
  Cell_t cell;
} __attribute__((packed));

struct LabelSst_t {
  Cell_t cell;
  uint32_t isst;
} __attribute__((packed));

struct Number_t {
  Cell_t cell;
  double num;
} __attribute__((packed));

struct RkNumber_t {
  uint32_t flags;

#define _nth_bit(_f, _nth) (((_f) & (1u << (_nth))) >> _nth)
  inline uint8_t fX100() const {
    return _nth_bit(flags, 0);
  }
  inline uint8_t fInt() const {
    return _nth_bit(flags, 1);
  }
  inline uint32_t num() const {
    return flags >> 2;
  }
#undef _nth_bit

  double value() const;
} __attribute__((packed));

struct RkRec_t {
  uint16_t ixfe;
  RkNumber_t RK;
} __attribute__((packed));

struct RK_t {
  uint16_t rw;
  uint16_t col;
  RkRec_t rkrec;
} __attribute__((packed));

class MulRk {
 public:
  int ParseFrom(const char *data, size_t data_len);

  inline uint16_t Rw() const {
    return m_rw;
  }
  inline uint16_t ColFirst() const {
    return m_colFirst;
  }
  inline uint16_t ColLast() const {
    return m_colLast;
  }
  inline const std::vector<RkRec_t> &RgRkrec() const {
    return m_rgrkrec;
  }

 private:
  uint16_t m_rw;
  uint16_t m_colFirst;
  uint16_t m_colLast;
  std::vector<RkRec_t> m_rgrkrec;
};

class MulBlank {
 public:
  int ParseFrom(const char *data, size_t data_len);

  inline uint16_t Rw() const {
    return m_rw;
  }
  inline uint16_t ColFirst() const {
    return m_colFirst;
  }
  inline uint16_t ColLast() const {
    return m_colLast;
  }

 private:
  uint16_t m_rw;
  uint16_t m_colFirst;
  uint16_t m_colLast;
};

class ShortXLUnicodeString {
 public:
  ssize_t ReadAndParse(const char *data, size_t data_len);

  inline const std::string &String() const {
    return m_str;
  }
  inline uint8_t Cch() const {
    return m_cch;
  }

 private:
  uint8_t m_cch;
  std::string m_str;
};

class XLUnicodeRichExtendedString {
 public:
  struct hdr_t {
    uint16_t cch;
    uint8_t flags1;

#define _nth_bit(_f, _nth) (((_f) & (1u << (_nth))) >> _nth)
    inline uint8_t fHighByte() const {
      return _nth_bit(flags1, 0);
    }
    inline uint8_t fExtSt() const {
      return _nth_bit(flags1, 2);
    }
    inline uint8_t fRichSt() const {
      return _nth_bit(flags1, 3);
    }
#undef _nth_bit
  } __attribute__((packed));

  ssize_t ReadAndParse(const char *data, size_t data_len, size_t offset,
                       size_t *curr_block_end);

  inline const hdr_t &Hdr() const {
    return m_hdr;
  }
  inline const std::string &String() const {
    return m_str;
  }

 private:
  hdr_t m_hdr;
  std::string m_str;
};

class BoundSheet8 {
 public:
  enum dt_t {
    kDT_WorksheetOrDialogSheet = 0x00,
    kDT_MacroSheet = 0x01,
    kDT_ChartSheet = 0x02,
    kDT_VbaModule = 0x06,
  };

  int ParseFrom(const char *data, size_t data_len);

  inline uint32_t LbPlyPos() const {
    return m_lbPlyPos;
  }
  inline uint8_t HsState() const {
    return m_hsState;
  }
  inline uint8_t Dt() const {
    return m_dt;
  }
  inline const ShortXLUnicodeString &Name() const {
    return m_name;
  }

 private:
  uint32_t m_lbPlyPos;
  uint8_t m_hsState;
  uint8_t m_dt;
  ShortXLUnicodeString m_name;
};

class MsXLS {
 public:
  int ParseFromFile(const std::string &filename);

  template <typename _T,
            typename std::enable_if<is_bytes<_T>::value>::type * = nullptr>
  int ParseFromBytes(_T &&data) {
    if (m_comp_doc.ParseFromBytes(std::forward<_T>(data)) != 0) {
      return -1;
    }
    return parse();
  }

  template <typename _T, typename std::enable_if<
                             is_compound_document<_T>::value>::type * = nullptr>
  int ParseFromCompoundDocument(_T &&d) {
    m_comp_doc = std::forward<_T>(d);
    return parse();
  }

  int FetchText(const fetch_text_options_t *opts, std::string *text) const;

 private:
  int parse();

 private:
  CompoundDocument m_comp_doc;
  int m_idx_workbook;

  // std::vector<char> m_workbook_stream;
  // std::vector<BoundSheet8> m_bs_list;
  // std::vector<XLUnicodeRichExtendedString> m_sst;
};

// =============================================================================

const std::string &Identifier2Name(uint16_t identifier);

ssize_t FetchTextFromSST(const record_header_t &rh, const char *data,
                         size_t data_len, size_t offset, int max_sst_cnt,
                         std::vector<XLUnicodeRichExtendedString> *sst);

ssize_t ReadAndParse1stSubstream(const char *data, size_t data_len,
                                 int max_sst_cnt,
                                 std::vector<BoundSheet8> *bs_list,
                                 std::vector<XLUnicodeRichExtendedString> *sst);

int Decrypt(char *data, size_t data_len,
            const std::wstring &password = L"VelvetSweatshop");

}  // namespace xls

}  // namespace msoffice
