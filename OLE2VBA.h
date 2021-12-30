//---------------------------------------------------------------------------
#ifndef __OLE2VBA_H__
#define __OLE2VBA_H__
//---------------------------------------------------------------------------
#include "OLE2Type.h"
#include "OLE2File.h"
#include "Streams.h"
//---------------------------------------------------------------------------
// Сигнатура потока VBA
#define VBA_STREAM_SIGNATURE    0x1601
// Идентификатор версии VBA96
#define VBA_VERSION_VBA96       0xB6
// Сигнатура заголовка VBA96
#define VBA96_HEADER_SIGNATURE  0x0001CAFE

// Заголовок потока VBA
#pragma pack(push, 1)
typedef struct _VBA_STREAM_HEADER
{
  uint16_t  wSign;                   // 0000: Сигнатура
  uint16_t  wUnknown1;               // 0002: ? (=1)
  uint8_t   bUnknown2;               // 0004: ?
  uint8_t   bVersion;                // 0005: Идентификатор версии
                                     //       (=0xB6 - VBA96)
  uint8_t   bUnknown3[23];           // 0006: ?
  uint32_t  dwNewVerMacroSrcOffset;  // 001D: Смещение исходного кода макроса
                                     //       для потока версии выше VBA96
} VBA_STREAM_HEADER, *PVBA_STREAM_HEADER;
#pragma pack(pop)
//---------------------------------------------------------------------------
// Извлечение исходного кода макроса из потока VBA
O2Res ExtractVBAMacroSource(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  );
// Детектирование потока VBA
bool DetectVBAStream(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID
  );
// Получение имени файла исходного кода макроса VBA
size_t MakeVBAMacroSourceFileName(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  wchar_t *pBuffer,
  size_t nBufferSize
  );
//---------------------------------------------------------------------------
#endif  // __OLE2VBA_H__
