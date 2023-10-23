#pragma once

#include <stdint.h>
#include <stdlib.h>

#include "msoffice/compound_document.h"

namespace msoffice {

namespace doc {

struct fib_base_t {
  uint16_t wIdent;
  uint16_t nFib;
  uint16_t unused;
  uint16_t lid;
  uint16_t pnNext;
  uint8_t flags1;
  uint8_t flags2;
  uint16_t nFibBack;
  uint32_t lKey;
  uint8_t envr;
  uint8_t flags3;
  uint16_t reserved3;
  uint16_t reserved4;
  uint32_t reserved5;
  uint32_t reserved6;

#define _nth_bit(_f, _nth) (((_f) & (1u << (_nth))) >> _nth)
  inline uint8_t fDot() const { return _nth_bit(flags1, 0); }
  inline uint8_t fGlsy() const { return _nth_bit(flags1, 1); }
  inline uint8_t fComplex() const { return _nth_bit(flags1, 2); }
  inline uint8_t fHasPic() const { return _nth_bit(flags1, 3); }
  inline uint8_t cQuickSaves() const { return flags1 >> 4; }
  inline uint8_t fEncrypted() const { return _nth_bit(flags2, 0); }
  inline uint8_t fWhichTblStm() const { return _nth_bit(flags2, 1); }
  inline uint8_t fReadOnlyRecommended() const { return _nth_bit(flags2, 2); }
  inline uint8_t fWriteReservation() const { return _nth_bit(flags2, 3); }
  inline uint8_t fExtChar() const { return _nth_bit(flags2, 4); }
  inline uint8_t fLoadOverride() const { return _nth_bit(flags2, 5); }
  inline uint8_t fFarEast() const { return _nth_bit(flags2, 6); }
  inline uint8_t fObfuscated() const { return flags2 >> 7; }
  inline uint8_t fMac() const { return _nth_bit(flags3, 0); }
  inline uint8_t fEmptySpecial() const { return _nth_bit(flags3, 1); }
  inline uint8_t fLoadOverridePage() const { return _nth_bit(flags3, 2); }
#undef _nth_bit
} __attribute__((packed));

struct FibRgFcLcb97_t {
  uint32_t fcStshfOrig;
  uint32_t lcbStshfOrig;
  uint32_t fcStshf;
  uint32_t lcbStshf;
  uint32_t fcPlcffndRef;
  uint32_t lcbPlcffndRef;
  uint32_t fcPlcffndTxt;
  uint32_t lcbPlcffndTxt;
  uint32_t fcPlcfandRef;
  uint32_t lcbPlcfandRef;
  uint32_t fcPlcfandTxt;
  uint32_t lcbPlcfandTxt;
  uint32_t fcPlcfSed;
  uint32_t lcbPlcfSed;
  uint32_t fcPlcPad;
  uint32_t lcbPlcPad;
  uint32_t fcPlcfPhe;
  uint32_t lcbPlcfPhe;
  uint32_t fcSttbfGlsy;
  uint32_t lcbSttbfGlsy;
  uint32_t fcPlcfGlsy;
  uint32_t lcbPlcfGlsy;
  uint32_t fcPlcfHdd;
  uint32_t lcbPlcfHdd;
  uint32_t fcPlcfBteChpx;
  uint32_t lcbPlcfBteChpx;
  uint32_t fcPlcfBtePapx;
  uint32_t lcbPlcfBtePapx;
  uint32_t fcPlcfSea;
  uint32_t lcbPlcfSea;
  uint32_t fcSttbfFfn;
  uint32_t lcbSttbfFfn;
  uint32_t fcPlcfFldMom;
  uint32_t lcbPlcfFldMom;
  uint32_t fcPlcfFldHdr;
  uint32_t lcbPlcfFldHdr;
  uint32_t fcPlcfFldFtn;
  uint32_t lcbPlcfFldFtn;
  uint32_t fcPlcfFldAtn;
  uint32_t lcbPlcfFldAtn;
  uint32_t fcPlcfFldMcr;
  uint32_t lcbPlcfFldMcr;
  uint32_t fcSttbfBkmk;
  uint32_t lcbSttbfBkmk;
  uint32_t fcPlcfBkf;
  uint32_t lcbPlcfBkf;
  uint32_t fcPlcfBkl;
  uint32_t lcbPlcfBkl;
  uint32_t fcCmds;
  uint32_t lcbCmds;
  uint32_t fcUnused1;
  uint32_t lcbUnused1;
  uint32_t fcSttbfMcr;
  uint32_t lcbSttbfMcr;
  uint32_t fcPrDrvr;
  uint32_t lcbPrDrvr;
  uint32_t fcPrEnvPort;
  uint32_t lcbPrEnvPort;
  uint32_t fcPrEnvLand;
  uint32_t lcbPrEnvLand;
  uint32_t fcWss;
  uint32_t lcbWss;
  uint32_t fcDop;
  uint32_t lcbDop;
  uint32_t fcSttbfAssoc;
  uint32_t lcbSttbfAssoc;
  uint32_t fcClx;
  uint32_t lcbClx;
  uint32_t fcPlcfPgdFtn;
  uint32_t lcbPlcfPgdFtn;
  uint32_t fcAutosaveSource;
  uint32_t lcbAutosaveSource;
  uint32_t fcGrpXstAtnOwners;
  uint32_t lcbGrpXstAtnOwners;
  uint32_t fcSttbfAtnBkmk;
  uint32_t lcbSttbfAtnBkmk;
  uint32_t fcUnused2;
  uint32_t lcbUnused2;
  uint32_t fcUnused3;
  uint32_t lcbUnused3;
  uint32_t fcPlcSpaMom;
  uint32_t lcbPlcSpaMom;
  uint32_t fcPlcSpaHdr;
  uint32_t lcbPlcSpaHdr;
  uint32_t fcPlcfAtnBkf;
  uint32_t lcbPlcfAtnBkf;
  uint32_t fcPlcfAtnBkl;
  uint32_t lcbPlcfAtnBkl;
  uint32_t fcPms;
  uint32_t lcbPms;
  uint32_t fcFormFldSttbs;
  uint32_t lcbFormFldSttbs;
  uint32_t fcPlcfendRef;
  uint32_t lcbPlcfendRef;
  uint32_t fcPlcfendTxt;
  uint32_t lcbPlcfendTxt;
  uint32_t fcPlcfFldEdn;
  uint32_t lcbPlcfFldEdn;
  uint32_t fcUnused4;
  uint32_t lcbUnused4;
  uint32_t fcDggInfo;
  uint32_t lcbDggInfo;
  uint32_t fcSttbfRMark;
  uint32_t lcbSttbfRMark;
  uint32_t fcSttbfCaption;
  uint32_t lcbSttbfCaption;
  uint32_t fcSttbfAutoCaption;
  uint32_t lcbSttbfAutoCaption;
  uint32_t fcPlcfWkb;
  uint32_t lcbPlcfWkb;
  uint32_t fcPlcfSpl;
  uint32_t lcbPlcfSpl;
  uint32_t fcPlcftxbxTxt;
  uint32_t lcbPlcftxbxTxt;
  uint32_t fcPlcfFldTxbx;
  uint32_t lcbPlcfFldTxbx;
  uint32_t fcPlcfHdrtxbxTxt;
  uint32_t lcbPlcfHdrtxbxTxt;
  uint32_t fcPlcffldHdrTxbx;
  uint32_t lcbPlcffldHdrTxbx;
  uint32_t fcStwUser;
  uint32_t lcbStwUser;
  uint32_t fcSttbTtmbd;
  uint32_t lcbSttbTtmbd;
  uint32_t fcCookieData;
  uint32_t lcbCookieData;
  uint32_t fcPgdMotherOldOld;
  uint32_t lcbPgdMotherOldOld;
  uint32_t fcBkdMotherOldOld;
  uint32_t lcbBkdMotherOldOld;
  uint32_t fcPgdFtnOldOld;
  uint32_t lcbPgdFtnOldOld;
  uint32_t fcBkdFtnOldOld;
  uint32_t lcbBkdFtnOldOld;
  uint32_t fcPgdEdnOldOld;
  uint32_t lcbPgdEdnOldOld;
  uint32_t fcBkdEdnOldOld;
  uint32_t lcbBkdEdnOldOld;
  uint32_t fcSttbfIntlFld;
  uint32_t lcbSttbfIntlFld;
  uint32_t fcRouteSlip;
  uint32_t lcbRouteSlip;
  uint32_t fcSttbSavedBy;
  uint32_t lcbSttbSavedBy;
  uint32_t fcSttbFnm;
  uint32_t lcbSttbFnm;
  uint32_t fcPlfLst;
  uint32_t lcbPlfLst;
  uint32_t fcPlfLfo;
  uint32_t lcbPlfLfo;
  uint32_t fcPlcfTxbxBkd;
  uint32_t lcbPlcfTxbxBkd;
  uint32_t fcPlcfTxbxHdrBkd;
  uint32_t lcbPlcfTxbxHdrBkd;
  uint32_t fcDocUndoWord9;
  uint32_t lcbDocUndoWord9;
  uint32_t fcRgbUse;
  uint32_t lcbRgbUse;
  uint32_t fcUsp;
  uint32_t lcbUsp;
  uint32_t fcUskf;
  uint32_t lcbUskf;
  uint32_t fcPlcupcRgbUse;
  uint32_t lcbPlcupcRgbUse;
  uint32_t fcPlcupcUsp;
  uint32_t lcbPlcupcUsp;
  uint32_t fcSttbGlsyStyle;
  uint32_t lcbSttbGlsyStyle;
  uint32_t fcPlgosl;
  uint32_t lcbPlgosl;
  uint32_t fcPlcocx;
  uint32_t lcbPlcocx;
  uint32_t fcPlcfBteLvc;
  uint32_t lcbPlcfBteLvc;
  uint32_t dwLowDateTime;
  uint32_t dwHighDateTime;
  uint32_t fcPlcfLvcPre10;
  uint32_t lcbPlcfLvcPre10;
  uint32_t fcPlcfAsumy;
  uint32_t lcbPlcfAsumy;
  uint32_t fcPlcfGram;
  uint32_t lcbPlcfGram;
  uint32_t fcSttbListNames;
  uint32_t lcbSttbListNames;
  uint32_t fcSttbfUssr;
  uint32_t lcbSttbfUssr;
} __attribute__((packed));

struct Pcd_FcCompressed_t {
  uint32_t flags;

#define _nth_bit(_f, _nth) (((_f) & (1u << (_nth))) >> _nth)
  inline uint32_t fc() const { return (flags << 2) >> 2; }
  inline uint32_t fCompressed() const { return _nth_bit(flags, 30); }
  inline uint32_t r1() const { return _nth_bit(flags, 31); }
#undef _nth_bit
} __attribute__((packed));

struct Pcd_Prm0_t {
  uint8_t flags1;
  uint8_t val;

  inline uint8_t fComplex() const { return flags1 & 1; }
  inline uint8_t isprm() const { return flags1 >> 1; }
} __attribute__((packed));

struct Pcd_t {
  uint16_t flags;
  Pcd_FcCompressed_t fc;
  Pcd_Prm0_t prm0;

#define _nth_bit(_f, _nth) (((_f) & (1u << (_nth))) >> _nth)
  inline uint16_t fNoParaLast() const { return _nth_bit(flags, 0); }
  inline uint16_t fR1() const { return _nth_bit(flags, 1); }
  inline uint16_t fDirty() const { return _nth_bit(flags, 2); }
  inline uint16_t fR2() const { return flags >> 3; }
#undef _nth_bit
} __attribute__((packed));

class PlcPcd {
 public:
  int ParseFrom(const FibRgFcLcb97_t& fib_rg_fc_lcb97,
                const std::vector<char>& table_stream);

  inline const std::vector<int32_t>& GetCP() const { return m_cp; }
  inline const std::vector<Pcd_t>& GetPcd() const { return m_pcd; }

 private:
  std::vector<int32_t> m_cp;
  std::vector<Pcd_t> m_pcd;
};

class MsDOC {
 public:
  int ParseFromFile(const std::string& filename);

  template <typename _T,
            typename std::enable_if<is_bytes<_T>::value>::type* = nullptr>
  int ParseFromBytes(_T&& data) {
    if (m_comp_doc.ParseFromBytes(std::forward<_T>(data)) != 0) {
      return -1;
    }
    return parse();
  }

  template <typename _T, typename std::enable_if<
                             is_compound_document<_T>::value>::type* = nullptr>
  int ParseFromCompoundDocument(_T&& d) {
    m_comp_doc = std::forward<_T>(d);
    return parse();
  }

  int FetchText(const fetch_text_options_t* opts, std::string* text) const;

 private:
  static const uint16_t _fib_base_wIdent;
  static const size_t _fib_rg_fc_lcb97_offset;

  int parse();

 private:
  CompoundDocument m_comp_doc;
  int m_idx_tab0;
  int m_idx_tab1;
  int m_idx_word_doc;
};

}  // namespace doc

}  // namespace msoffice
