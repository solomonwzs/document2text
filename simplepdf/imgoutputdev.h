#pragma once

#include <poppler/OutputDev.h>

namespace simplepdf {

class ImageOutputDev : public OutputDev {
 public:
  ImageOutputDev();

  bool upsideDown() override { return true; }
  bool useDrawChar() override { return false; }
  bool interpretType3Chars() override { return false; }
  bool needNonText() override { return true; }

  void drawImageMask(GfxState *state, Object *ref, Stream *str, int width,
                     int height, bool invert, bool interpolate,
                     bool inlineImg) override;
  void drawImage(GfxState *state, Object *ref, Stream *str, int width,
                 int height, GfxImageColorMap *colorMap, bool interpolate,
                 const int *maskColors, bool inlineImg) override;
  void drawMaskedImage(GfxState *state, Object *ref, Stream *str, int width,
                       int height, GfxImageColorMap *colorMap, bool interpolate,
                       Stream *maskStr, int maskWidth, int maskHeight,
                       bool maskInvert, bool maskInterpolate) override;
  void drawSoftMaskedImage(GfxState *state, Object *ref, Stream *str, int width,
                           int height, GfxImageColorMap *colorMap,
                           bool interpolate, Stream *maskStr, int maskWidth,
                           int maskHeight, GfxImageColorMap *maskColorMap,
                           bool maskInterpolate) override;

 private:
  int m_cur_page;
};

}  // namespace simplepdf
