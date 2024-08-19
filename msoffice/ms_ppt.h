#pragma once

#include <map>
#include <string>

#include "msoffice/compound_document.h"
#include "msoffice/utils.h"

namespace msoffice {

namespace ppt {

struct record_header_t {
  uint16_t flags;
  uint16_t recType;
  uint32_t recLen;

  inline uint16_t recVer() const {
    return flags & 0b1111;
  }
  inline uint16_t recInstance() const {
    return flags >> 4;
  }
} __attribute__((packed));

enum record_type_t {
  kRT_Document = 0x03E8,
  kRT_DocumentAtom = 0x03E9,
  kRT_EndDocumentAtom = 0x03EA,
  kRT_Slide = 0x03EE,
  kRT_SlideAtom = 0x03EF,
  kRT_Notes = 0x03F0,
  kRT_NotesAtom = 0x03F1,
  kRT_Environment = 0x03F2,
  kRT_SlidePersistAtom = 0x03F3,
  kRT_MainMaster = 0x03F8,
  kRT_SlideShowSlideInfoAtom = 0x03F9,
  kRT_SlideViewInfo = 0x03FA,
  kRT_GuideAtom = 0x03FB,
  kRT_ViewInfoAtom = 0x03FD,
  kRT_SlideViewInfoAtom = 0x03FE,
  kRT_VbaInfo = 0x03FF,
  kRT_VbaInfoAtom = 0x0400,
  kRT_SlideShowDocInfoAtom = 0x0401,
  kRT_Summary = 0x0402,
  kRT_DocRoutingSlipAtom = 0x0406,
  kRT_OutlineViewInfo = 0x0407,
  kRT_SorterViewInfo = 0x0408,
  kRT_ExternalObjectList = 0x0409,
  kRT_ExternalObjectListAtom = 0x040A,
  kRT_DrawingGroup = 0x040B,
  kRT_Drawing = 0x040C,
  kRT_GridSpacing10Atom = 0x040D,
  kRT_RoundTripTheme12Atom = 0x040E,
  kRT_RoundTripColorMapping12Atom = 0x040F,
  kRT_NamedShows = 0x0410,
  kRT_NamedShow = 0x0411,
  kRT_NamedShowSlidesAtom = 0x0412,
  kRT_NotesTextViewInfo9 = 0x0413,
  kRT_NormalViewSetInfo9 = 0x0414,
  kRT_NormalViewSetInfo9Atom = 0x0415,
  kRT_RoundTripOriginalMainMasterId12Atom = 0x041C,
  kRT_RoundTripCompositeMasterId12Atom = 0x041D,
  kRT_RoundTripContentMasterInfo12Atom = 0x041E,
  kRT_RoundTripShapeId12Atom = 0x041F,
  kRT_RoundTripHFPlaceholder12Atom = 0x0420,
  kRT_RoundTripContentMasterId12Atom = 0x0422,
  kRT_RoundTripOArtTextStyles12Atom = 0x0423,
  kRT_RoundTripHeaderFooterDefaults12Atom = 0x0424,
  kRT_RoundTripDocFlags12Atom = 0x0425,
  kRT_RoundTripShapeCheckSumForCL12Atom = 0x0426,
  kRT_RoundTripNotesMasterTextStyles12Atom = 0x0427,
  kRT_RoundTripCustomTableStyles12Atom = 0x0428,
  kRT_List = 0x07D0,
  kRT_FontCollection = 0x07D5,
  kRT_FontCollection10 = 0x07D6,
  kRT_BookmarkCollection = 0x07E3,
  kRT_SoundCollection = 0x07E4,
  kRT_SoundCollectionAtom = 0x07E5,
  kRT_Sound = 0x07E6,
  kRT_SoundDataBlob = 0x07E7,
  kRT_BookmarkSeedAtom = 0x07E9,
  kRT_ColorSchemeAtom = 0x07F0,
  kRT_BlipCollection9 = 0x07F8,
  kRT_BlipEntity9Atom = 0x07F9,
  kRT_ExternalObjectRefAtom = 0x0BC1,
  kRT_PlaceholderAtom = 0x0BC3,
  kRT_ShapeAtom = 0x0BDB,
  kRT_ShapeFlags10Atom = 0x0BDC,
  kRT_RoundTripNewPlaceholderId12Atom = 0x0BDD,
  kRT_OutlineTextRefAtom = 0x0F9E,
  kRT_TextHeaderAtom = 0x0F9F,
  kRT_TextCharsAtom = 0x0FA0,
  kRT_StyleTextPropAtom = 0x0FA1,
  kRT_MasterTextPropAtom = 0x0FA2,
  kRT_TextMasterStyleAtom = 0x0FA3,
  kRT_TextCharFormatExceptionAtom = 0x0FA4,
  kRT_TextParagraphFormatExceptionAtom = 0x0FA5,
  kRT_TextRulerAtom = 0x0FA6,
  kRT_TextBookmarkAtom = 0x0FA7,
  kRT_TextBytesAtom = 0x0FA8,
  kRT_TextSpecialInfoDefaultAtom = 0x0FA9,
  kRT_TextSpecialInfoAtom = 0x0FAA,
  kRT_DefaultRulerAtom = 0x0FAB,
  kRT_StyleTextProp9Atom = 0x0FAC,
  kRT_TextMasterStyle9Atom = 0x0FAD,
  kRT_OutlineTextProps9 = 0x0FAE,
  kRT_OutlineTextPropsHeader9Atom = 0x0FAF,
  kRT_TextDefaults9Atom = 0x0FB0,
  kRT_StyleTextProp10Atom = 0x0FB1,
  kRT_TextMasterStyle10Atom = 0x0FB2,
  kRT_OutlineTextProps10 = 0x0FB3,
  kRT_TextDefaults10Atom = 0x0FB4,
  kRT_OutlineTextProps11 = 0x0FB5,
  kRT_StyleTextProp11Atom = 0x0FB6,
  kRT_FontEntityAtom = 0x0FB7,
  kRT_FontEmbedDataBlob = 0x0FB8,
  kRT_CString = 0x0FBA,
  kRT_MetaFile = 0x0FC1,
  kRT_ExternalOleObjectAtom = 0x0FC3,
  kRT_Kinsoku = 0x0FC8,
  kRT_Handout = 0x0FC9,
  kRT_ExternalOleEmbed = 0x0FCC,
  kRT_ExternalOleEmbedAtom = 0x0FCD,
  kRT_ExternalOleLink = 0x0FCE,
  kRT_BookmarkEntityAtom = 0x0FD0,
  kRT_ExternalOleLinkAtom = 0x0FD1,
  kRT_KinsokuAtom = 0x0FD2,
  kRT_ExternalHyperlinkAtom = 0x0FD3,
  kRT_ExternalHyperlink = 0x0FD7,
  kRT_SlideNumberMetaCharAtom = 0x0FD8,
  kRT_HeadersFooters = 0x0FD9,
  kRT_HeadersFootersAtom = 0x0FDA,
  kRT_TextInteractiveInfoAtom = 0x0FDF,
  kRT_ExternalHyperlink9 = 0x0FE4,
  kRT_RecolorInfoAtom = 0x0FE7,
  kRT_ExternalOleControl = 0x0FEE,
  kRT_SlideListWithText = 0x0FF0,
  kRT_AnimationInfoAtom = 0x0FF1,
  kRT_InteractiveInfo = 0x0FF2,
  kRT_InteractiveInfoAtom = 0x0FF3,
  kRT_UserEditAtom = 0x0FF5,
  kRT_CurrentUserAtom = 0x0FF6,
  kRT_DateTimeMetaCharAtom = 0x0FF7,
  kRT_GenericDateMetaCharAtom = 0x0FF8,
  kRT_HeaderMetaCharAtom = 0x0FF9,
  kRT_FooterMetaCharAtom = 0x0FFA,
  kRT_ExternalOleControlAtom = 0x0FFB,
  kRT_ExternalMediaAtom = 0x1004,
  kRT_ExternalVideo = 0x1005,
  kRT_ExternalAviMovie = 0x1006,
  kRT_ExternalMciMovie = 0x1007,
  kRT_ExternalMidiAudio = 0x100D,
  kRT_ExternalCdAudio = 0x100E,
  kRT_ExternalWavAudioEmbedded = 0x100F,
  kRT_ExternalWavAudioLink = 0x1010,
  kRT_ExternalOleObjectStg = 0x1011,
  kRT_ExternalCdAudioAtom = 0x1012,
  kRT_ExternalWavAudioEmbeddedAtom = 0x1013,
  kRT_AnimationInfo = 0x1014,
  kRT_RtfDateTimeMetaCharAtom = 0x1015,
  kRT_ExternalHyperlinkFlagsAtom = 0x1018,
  kRT_ProgTags = 0x1388,
  kRT_ProgStringTag = 0x1389,
  kRT_ProgBinaryTag = 0x138A,
  kRT_BinaryTagDataBlob = 0x138B,
  kRT_PrintOptionsAtom = 0x1770,
  kRT_PersistDirectoryAtom = 0x1772,
  kRT_PresentationAdvisorFlags9Atom = 0x177A,
  kRT_HtmlDocInfo9Atom = 0x177B,
  kRT_HtmlPublishInfoAtom = 0x177C,
  kRT_HtmlPublishInfo9 = 0x177D,
  kRT_BroadcastDocInfo9 = 0x177E,
  kRT_BroadcastDocInfo9Atom = 0x177F,
  kRT_EnvelopeFlags9Atom = 0x1784,
  kRT_EnvelopeData9Atom = 0x1785,
  kRT_VisualShapeAtom = 0x2AFB,
  kRT_HashCodeAtom = 0x2B00,
  kRT_VisualPageAtom = 0x2B01,
  kRT_BuildList = 0x2B02,
  kRT_BuildAtom = 0x2B03,
  kRT_ChartBuild = 0x2B04,
  kRT_ChartBuildAtom = 0x2B05,
  kRT_DiagramBuild = 0x2B06,
  kRT_DiagramBuildAtom = 0x2B07,
  kRT_ParaBuild = 0x2B08,
  kRT_ParaBuildAtom = 0x2B09,
  kRT_LevelInfoAtom = 0x2B0A,
  kRT_RoundTripAnimationAtom12Atom = 0x2B0B,
  kRT_RoundTripAnimationHashAtom12Atom = 0x2B0D,
  kRT_Comment10 = 0x2EE0,
  kRT_Comment10Atom = 0x2EE1,
  kRT_CommentIndex10 = 0x2EE4,
  kRT_CommentIndex10Atom = 0x2EE5,
  kRT_LinkedShape10Atom = 0x2EE6,
  kRT_LinkedSlide10Atom = 0x2EE7,
  kRT_SlideFlags10Atom = 0x2EEA,
  kRT_SlideTime10Atom = 0x2EEB,
  kRT_DiffTree10 = 0x2EEC,
  kRT_Diff10 = 0x2EED,
  kRT_Diff10Atom = 0x2EEE,
  kRT_SlideListTableSize10Atom = 0x2EEF,
  kRT_SlideListEntry10Atom = 0x2EF0,
  kRT_SlideListTable10 = 0x2EF1,
  kRT_CryptSession10Container = 0x2F14,
  kRT_FontEmbedFlags10Atom = 0x32C8,
  kRT_FilterPrivacyFlags10Atom = 0x36B0,
  kRT_DocToolbarStates10Atom = 0x36B1,
  kRT_PhotoAlbumInfo10Atom = 0x36B2,
  kRT_SmartTagStore11Container = 0x36B3,
  kRT_RoundTripSlideSyncInfo12 = 0x3714,
  kRT_RoundTripSlideSyncInfoAtom12 = 0x3715,
  kRT_TimeConditionContainer = 0xF125,
  kRT_TimeNode = 0xF127,
  kRT_TimeCondition = 0xF128,
  kRT_TimeModifier = 0xF129,
  kRT_TimeBehaviorContainer = 0xF12A,
  kRT_TimeAnimateBehaviorContainer = 0xF12B,
  kRT_TimeColorBehaviorContainer = 0xF12C,
  kRT_TimeEffectBehaviorContainer = 0xF12D,
  kRT_TimeMotionBehaviorContainer = 0xF12E,
  kRT_TimeRotationBehaviorContainer = 0xF12F,
  kRT_TimeScaleBehaviorContainer = 0xF130,
  kRT_TimeSetBehaviorContainer = 0xF131,
  kRT_TimeCommandBehaviorContainer = 0xF132,
  kRT_TimeBehavior = 0xF133,
  kRT_TimeAnimateBehavior = 0xF134,
  kRT_TimeColorBehavior = 0xF135,
  kRT_TimeEffectBehavior = 0xF136,
  kRT_TimeMotionBehavior = 0xF137,
  kRT_TimeRotationBehavior = 0xF138,
  kRT_TimeScaleBehavior = 0xF139,
  kRT_TimeSetBehavior = 0xF13A,
  kRT_TimeCommandBehavior = 0xF13B,
  kRT_TimeClientVisualElement = 0xF13C,
  kRT_TimePropertyList = 0xF13D,
  kRT_TimeVariantList = 0xF13E,
  kRT_TimeAnimationValueList = 0xF13F,
  kRT_TimeIterateData = 0xF140,
  kRT_TimeSequenceData = 0xF141,
  kRT_TimeVariant = 0xF142,
  kRT_TimeAnimationValue = 0xF143,
  kRT_TimeExtTimeNodeContainer = 0xF144,
  kRT_TimeSubEffectContainer = 0xF145,

  kRT_OfficeArtDg = 0xF002,
  kRT_OfficeArtFDG = 0xF008,
  kRT_OfficeArtSpgrContainer = 0xF003,
  kRT_OfficeArtSpContainer = 0xF004,
  kRT_OfficeArtClientTextbox = 0xF00D,
};

class CurrentUserAtom {
 public:
  struct hdr_t {
    record_header_t rh;
    uint32_t size;
    uint32_t headerToken;
    uint32_t offsetToCurrentEdit;
    uint16_t lenUserName;
    uint16_t docFileVersion;
    uint8_t majorVersion;
    uint8_t minorVersion;
    uint16_t unused;
  } __attribute__((packed));

 public:
  ssize_t ReadAndParse(const char *data, size_t size);

  inline const hdr_t &Header() const {
    return m_hdr;
  }
  inline const std::string UserName() const {
    return m_user_name;
  }
  inline bool IsEncrypted() const {
    return m_hdr.headerToken == 0xF3D1C4DF;
  }

 private:
  hdr_t m_hdr;
  uint32_t m_rel_version;
  std::string m_ansi_user_name;
  std::string m_user_name;
};

struct UserEditAtom_t {
  record_header_t rh;
  uint32_t lastSlideIdRef;
  uint16_t version;
  uint8_t minorVersion;
  uint8_t majorVersion;
  uint32_t offsetLastEdit;
  uint32_t offsetPersistDirectory;
  uint32_t docPersistIdRef;
  uint32_t persistIdSeed;
  uint16_t lastView;
  uint16_t unused;
  // uint32_t encryptSessionPersistIdRef;
} __attribute__((packed));

struct PersistDirectoryEntry_t {
  uint32_t flags;
  uint32_t persistOffset[];

  inline uint32_t persistId() const {
    return flags & ((1u << 20) - 1u);
  }
  inline uint16_t cPersist() const {
    return flags >> 20;
  }
} __attribute__((packed));

class PersistDirectoryAtom {
 public:
  ssize_t ReadAndParse(const UserEditAtom_t &user_edit_atom, const char *data,
                       size_t data_len);

  static int ParsePersistDirectoryEntry(
      const UserEditAtom_t &user_edit_atom, const char *data, size_t data_len,
      std::vector<const PersistDirectoryEntry_t *> *entries);

  inline const record_header_t &RecordHeader() const {
    return m_rh;
  }
  inline const std::vector<const PersistDirectoryEntry_t *> &Entries() const {
    return m_persistDirectoryEntry;
  }

 private:
  PersistDirectoryAtom() = default;

 private:
  record_header_t m_rh;
  std::vector<char> m_data;
  std::vector<const PersistDirectoryEntry_t *> m_persistDirectoryEntry;
};

class MsPPT {
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

  static int GetPersistId2Offset(const CurrentUserAtom &current_user_atom,
                                 const char *ppt_doc_stream,
                                 size_t ppt_doc_stream_len,
                                 std::map<uint32_t, uint32_t> *id2offset);

  int FetchText(const fetch_text_options_t *opts, std::string *text) const;

  int parse();

 private:
  CompoundDocument m_comp_doc;
  int m_idx_current_user;
  int m_idx_ppt_doc;
};

}  // namespace ppt

}  // namespace msoffice
