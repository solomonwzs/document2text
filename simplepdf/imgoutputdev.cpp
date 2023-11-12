#include "simplepdf/imgoutputdev.h"

#include <stdio.h>

#define _STYLE_Debug "\e[3;36m"
#define _STYLE_Info  "\e[3;32m"
#define _STYLE_Warn  "\e[3;33m"
#define _STYLE_Err   "\e[3;31m"
#define slog(_type_, _fmt_, ...)                                      \
  printf(_STYLE_##_type_ "%.1s [%s:%s:%d]\e[0m " _fmt_ "\n", #_type_, \
         __FILE__, __func__, __LINE__, ##__VA_ARGS__)

namespace simplepdf {

ImageOutputDev::ImageOutputDev() : m_cur_page(0) {}

void ImageOutputDev::drawImageMask(GfxState *state, Object *ref, Stream *str,
                                   int width, int height, bool invert,
                                   bool interpolate, bool inlineImg) {
  slog(Debug, "%d %d", width, height);
}

void ImageOutputDev::drawImage(GfxState *state, Object *ref, Stream *str,
                               int width, int height,
                               GfxImageColorMap *colorMap, bool interpolate,
                               const int *maskColors, bool inlineImg) {
  slog(Debug, "%d %d", width, height);

  str = str->getNextStream();
  str->reset();

  int c;
  size_t len = 0;
  while ((c = str->getChar()) != EOF) {
    len += 1;
  }
  str->close();
  slog(Debug, "%ld", len);
}

void ImageOutputDev::drawMaskedImage(GfxState *state, Object *ref, Stream *str,
                                     int width, int height,
                                     GfxImageColorMap *colorMap,
                                     bool interpolate, Stream *maskStr,
                                     int maskWidth, int maskHeight,
                                     bool maskInvert, bool maskInterpolate) {
  slog(Debug, "%d %d", width, height);
}

void ImageOutputDev::drawSoftMaskedImage(GfxState *state, Object *ref,
                                         Stream *str, int width, int height,
                                         GfxImageColorMap *colorMap,
                                         bool interpolate, Stream *maskStr,
                                         int maskWidth, int maskHeight,
                                         GfxImageColorMap *maskColorMap,
                                         bool maskInterpolate) {
  slog(Debug, "%d %d", width, height);
}

}  // namespace simplepdf
