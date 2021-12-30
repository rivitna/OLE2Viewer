//---------------------------------------------------------------------------
#ifdef USE_WINAPI
#define WIN32_LEAN_AND_MEAN  // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#else
#include <string.h>
#endif  // USE_WINAPI

#include <malloc.h>
#include <memory.h>
#include "StrUtils.h"
#include "OLE2CF.h"
#include "OLE2VBA.h"
//---------------------------------------------------------------------------
// Наименование каталога VBA
const wchar_t VBA_STORAGE_NAME[]     = L"VBA";
// Расширение файла с исходным кодом макроса VBA
const wchar_t VBA_MACRO_FILE_EXT[]   = L".bas";
// Имя модуля VBA по умолчанию
const wchar_t VBA_MODULE_FILE_NAME[] = L"Module";
//---------------------------------------------------------------------------
/***************************************************************************/
/* FindVBAMacroSource - Поиск запакованного исходного кода макроса         */
/***************************************************************************/
const void* FindVBAMacroSource(
  const void *buf,
  size_t count
  )
{
  if (count < 2 * sizeof(uint32_t))
    return NULL;

  count -= 2 * sizeof(uint32_t) - 1;
  do
  {
    if (((((uint32_t *)buf)[0] & 0xFF7000FF) == 0x00300001) &&
        (((uint32_t *)buf)[1] == 0x72747441))  /* "Attr" */
      return buf;
    buf = (uint8_t *)buf + 1;
    count--;
  }
  while (count != 0);
  return NULL;
}
/***************************************************************************/
/* SearchUInt32 - Поиск 32-разрядного числа в блоке памяти                 */
/***************************************************************************/
const void *SearchUInt32(
  const void *buf,
  size_t count,
  uint32_t val
  )
{
  if (count < sizeof(uint32_t))
    return NULL;

  count -= sizeof(uint32_t) - 1;
  do
  {
    if (val == ((uint32_t *)buf)[0])
      return buf;
    buf = (uint8_t *)buf + 1;
    count--;
  }
  while (count != 0);
  return NULL;
}
/***************************************************************************/
/* GetVBA96MacroSourcePtr - Получение адреса исходного кода макроса        */
/*                          в потоке версии VBA96                          */
/***************************************************************************/
const void* GetVBA96MacroSourcePtr(
  const void *pStreamData,
  size_t cbStreamData
  )
{
  const uint8_t *pb;
  size_t cb;
  size_t nOffset;

  if (cbStreamData <= sizeof_through_field(VBA_STREAM_HEADER, bVersion) +
                      sizeof(uint32_t))
    return NULL;

  // Поиск сигнатуры заголовка VBA96
  pb = (uint8_t *)SearchUInt32((uint8_t *)pStreamData +
                               sizeof_through_field(VBA_STREAM_HEADER,
                                                    bVersion),
                               cbStreamData -
                               sizeof_through_field(VBA_STREAM_HEADER,
                                                    bVersion),
                               VBA96_HEADER_SIGNATURE);
  if (!pb)
    return NULL;
  cb = cbStreamData - (pb - (uint8_t *)pStreamData);
  if (cb < sizeof(uint32_t) + sizeof(uint16_t))
    return NULL;
  nOffset = (((uint16_t *)(pb + sizeof(uint32_t)))[0] + 1) * 12;
  if (nOffset >= cb)
    return NULL;
  pb += nOffset;
  cb -= nOffset;
  if (cb < sizeof(uint32_t))
    return NULL;
  nOffset = ((uint32_t *)pb)[0] + 4;
  if (nOffset >= cb)
    return NULL;
  return (pb + nOffset);
}
/***************************************************************************/
/* LZNT1_Decompress - Распаковка данных, сжатых по алгоритму LZNT1         */
/***************************************************************************/
#ifndef _WIN64
__declspec(naked)
#endif  // _WIN64
size_t
#ifndef _WIN64
__stdcall
#endif  // _WIN64
LZNT1_Decompress(
  void *dest,
  size_t destlen,
  const void *src,
  size_t srclen
  )
{
#ifdef _WIN64
  union
  {
    const uint8_t  *pb;
    const uint16_t *pw;
  } src_ptr;
  uint8_t *pd;
  unsigned int chunk_len;
  unsigned int chunk_written;
  unsigned int next_threshold;
  unsigned int split;
  unsigned int tag_bit;
  unsigned int tag;
  unsigned int copy_len;
  uint8_t *copy_src;

  if (!src || (srclen == 0))
    return 0;

  if (!dest)
    destlen = 0;

  src_ptr.pb = (const uint8_t *)src;
  pd = (uint8_t *)dest;

  while (srclen >= sizeof(uint16_t))
  {
    // Проверка сигнатуры блока
    if (((*src_ptr.pw >> 8) & 0x70) != 0x30)
      break;

    // Определение длины данных блока
    chunk_len = (*src_ptr.pw & 0x0FFF) + 1;
    srclen -= sizeof(uint16_t);
    if (srclen < chunk_len)
      chunk_len = (unsigned int)srclen;
    srclen -= chunk_len;

    if ((*(src_ptr.pw++) >> 8) & 0x80)
    {
      // Сжатый блок
      chunk_written = 0;
      next_threshold = 16;
      split = 12;
      tag_bit = 0;

      while (chunk_len != 0)
      {
        if (tag_bit == 0)
        {
          tag = *(src_ptr.pb++);
          chunk_len--;
          if (chunk_len == 0)
            break;
        }

        chunk_len--;

        if (tag & 1)
        {
          if (chunk_len == 0)
            break;
          while (chunk_written > next_threshold)
          {
            split--;
            next_threshold <<= 1;
          }
          copy_len = (*src_ptr.pw & ((1 << split) - 1)) + 3;
          copy_src = pd - ((*src_ptr.pw >> split) + 1);
          src_ptr.pw++;
          chunk_len--;
          chunk_written += copy_len;
          while ((copy_len != 0) && (destlen != 0))
          {
            *(pd++) = *(copy_src++);
            copy_len--;
            destlen--;
          }
          pd += copy_len;
        }
        else
        {
          if (destlen != 0)
          {
            *pd = *src_ptr.pb;
            destlen--;
          }
          pd++;
          src_ptr.pb++;
          chunk_written++;
        }

        tag >>= 1;
        tag_bit = (tag_bit + 1) & 7;
      }
    }
    else
    {
      // Несжатый блок
      while ((chunk_len != 0) && (destlen != 0))
      {
        *(pd++) = *(src_ptr.pb++);
        chunk_len--;
        destlen--;
      }
      pd += chunk_len;
    }
  }

  return (pd - (uint8_t *)dest);
#else
  __asm
  {
    push  ebp
    push  ebx
    push  esi
    push  edi

    xor   eax,eax          // EAX=0

    mov   esi,[esp+1Ch]    // ESI = src
    or    esi,esi
    jz    ToEnd
    cmp   eax,[esp+20h]    // srclen != 0?
    je    ToEnd

    xor   ebx,ebx
    mov   edi,[esp+14h]    // EDI = dest
    or    edi,edi
    jz    ChunkLoop
    mov   ebx,[esp+18h]    // EBX = destlen

ChunkLoop:
    mov   edx,[esp+20h]    // EDX = текущий размер входного буфера
    sub   edx,2
    jb    Done

    lodsw                  // AX = заголовок блока
    mov   ecx,eax          // ECX = заголовок блока
    mov   al,ah            // AL = сигнатура блока
    and   al,070h
    cmp   al,030h          // Проверка сигнатуры блока
    je    DecompressChunk

Done:
    mov   eax,edi
    sub   eax,[esp+14h]    // EAX = число записанных в буфер байтов

ToEnd:
    pop   edi
    pop   esi
    pop   ebx
    pop   ebp
    ret   16

DecompressChunk:
    and   cx,0FFFh
    inc   ecx              // ECX = длина данных блока

    cmp   ecx,edx
    jbe   UpdateSrcLen
    mov   ecx,edx
    jecxz ChunkLoop
UpdateSrcLen:
    sub   edx,ecx
    mov   [esp+20h],edx

    test  ah,ah            // Сжатый блок?
    js    CompressedChunk

    or    ebx,ebx
    jz    CopyLoopEnd1
CopyLoop1:
    movsb
    dec   ebx
    loopnz CopyLoop1
CopyLoopEnd1:
    add   edi,ecx
    jmp   ChunkLoop

CompressedChunk:
    lea   ebp,[ecx+1]      // EBP = длина блока + 1
    cdq
    dec   edx              // EDX=-1

ByteLoop:
    dec   ebp
    jz    ChunkLoop

    lodsb
    mov   ch,al
    mov   cl,8

BitLoop:
    dec   ebp
    jz    ChunkLoop

    shr   ch,1
    jc    Match

    or    ebx,ebx
    jz    CopyByteEnd
    mov   al,[esi]
    mov   [edi],al
    dec   ebx
CopyByteEnd:
    inc   esi
    inc   edi
    inc   edx
    jmp   NextBit

Match:
    dec   ebp
    jz    ChunkLoop

    lodsw

    push  ecx
    push  esi

    mov   ecx,edx          // ECX = количество записанных в буфер байтов для
                           //       текущего блока - 1
    or    cl,8
    bsr   cx,cx
    xor   cl,0Fh           // CL = количество битов для сдвига
    xor   esi,esi
    inc   esi
    shl   esi,cl
    dec   esi              // ESI = маска количества копируемых байтов
    and   esi,eax
    add   esi,3            // ESI = число копируемых байтов
    add   edx,esi          // EDX = количество записанных в буфер байтов для
                           //       текущего блока - 1
    shr   eax,cl           // EAX = смещение копируемой цепочки байтов - 1
    mov   ecx,esi          // ECX = число копируемых байтов

    or    ebx,ebx
    jz    CopyLoopEnd2
    mov   esi,edi
    stc
    sbb   esi,eax          // ESI -> копируемая цепочка байтов
CopyLoop2:
    movsb
    dec   ebx
    loopnz CopyLoop2
CopyLoopEnd2:
    add   edi,ecx

    pop   esi
    pop   ecx

NextBit:
    dec   cl
    jnz   BitLoop
    jmp   ByteLoop
  }
#endif  // _WIN64
}
/***************************************************************************/
/* DecompressVBAMacroSource - Распаковка исходного кода макроса VBA        */
/***************************************************************************/
O2Res DecompressVBAMacroSource(
  const void *pPackedMacroSrc,
  size_t cbPackedMacroSrc,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  )
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  // Проверка маркера запакованного исходного кода макроса VBA
  if ((cbPackedMacroSrc <= 1) ||
      (((uint8_t *)pPackedMacroSrc)[0] != 0x01))
    return O2_ERROR_DATA;

  pPackedMacroSrc = (uint8_t *)pPackedMacroSrc + 1;
  cbPackedMacroSrc--;

  size_t cbMacroSrc;

  // Определение размера распакованного исходного кода макроса VBA
  cbMacroSrc = LZNT1_Decompress(NULL, 0, pPackedMacroSrc, cbPackedMacroSrc);
  if (cbMacroSrc == 0)
    return O2_ERROR_DATA;

  O2Res res = O2_OK;

  if (pOutStream)
  {
    // Выделение памяти для исходного кода макроса VBA
    void *pMacroSrc = malloc(cbMacroSrc);
    if (!pMacroSrc)
      return O2_ERROR_MEM;

    // Распаковка данных, сжатых по алгоритму LZNT1
    cbMacroSrc = LZNT1_Decompress(pMacroSrc, cbMacroSrc,
                                  pPackedMacroSrc, cbPackedMacroSrc);

    // Запись данных в поток
    size_t nBytesWritten = pOutStream->Write(pMacroSrc, cbMacroSrc);
    if (pnProcessedSize)
      *pnProcessedSize = nBytesWritten;
    if (nBytesWritten != cbMacroSrc)
      res = O2_ERROR_WRITE;

    free(pMacroSrc);
  }
  else
  {
    if (pnProcessedSize)
      *pnProcessedSize = cbMacroSrc;
  }

  return res;
}
/***************************************************************************/
/* ExtractVBAMacroSource - Извлечение исходного кода макроса из потока VBA */
/***************************************************************************/
O2Res ExtractVBAMacroSource(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  )
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  if (!pOLE2File || !pOLE2File->GetOpened() ||
      (dwStreamID == NOSTREAM))
    return O2_ERROR_PARAM;

  const DIRECTORY_ENTRY *pDirEntry;
  pDirEntry = pOLE2File->GetDirEntry(dwStreamID);
  if (!pDirEntry ||
      (pDirEntry->bObjectType != OBJ_TYPE_STREAM))
    return O2_ERROR_PARAM;

  uint64_t nStreamSize = (pOLE2File->GetMajorVersion() > 3)
                           ? pDirEntry->qwStreamSize
                           : pDirEntry->dwLowStreamSize;
  if ((nStreamSize > V3_MAX_STREAM_SIZE) ||
      (pDirEntry->dwLowStreamSize <= sizeof_through_field(VBA_STREAM_HEADER,
                                                          bVersion)))
    return O2_ERROR_PARAM;

  O2Res res;
  size_t nBytesRead;
  VBA_STREAM_HEADER vbaStreamHdr;

  // Чтение заголовка потока VBA
  res = pOLE2File->ExtractStreamData(dwStreamID, 0, sizeof(vbaStreamHdr),
                                     &vbaStreamHdr, &nBytesRead);
  if ((nBytesRead < sizeof(vbaStreamHdr.wSign)) ||
      (vbaStreamHdr.wSign != VBA_STREAM_SIGNATURE))
    return (res != O2_OK) ? res : O2_ERROR_SIGN;
  if (nBytesRead <= sizeof_through_field(VBA_STREAM_HEADER, bVersion))
    return (res != O2_OK) ? res : O2_ERROR_DATA;

  bool bVBA96 = (vbaStreamHdr.bVersion == VBA_VERSION_VBA96);
  size_t nDataOffset = 0;
  size_t nDataSize = pDirEntry->dwLowStreamSize;

  if (!bVBA96)
  {
    // Получение для потока версии выше VBA96 смещения и размера
    // запакованного исходного кода макроса
    if ((nBytesRead < sizeof_through_field(VBA_STREAM_HEADER,
                                           dwNewVerMacroSrcOffset)) ||
        (vbaStreamHdr.dwNewVerMacroSrcOffset < sizeof(VBA_STREAM_HEADER)) ||
        (vbaStreamHdr.dwNewVerMacroSrcOffset >= nDataSize - 6))
      return (res != O2_OK) ? res : O2_ERROR_DATA;
    nDataOffset = vbaStreamHdr.dwNewVerMacroSrcOffset + 6;
    nDataSize -= nDataOffset;
  }

  if (nDataSize < sizeof(uint32_t))
    return O2_ERROR_DATA;

  // Выделение памяти для данных потока
  uint8_t *pbData = (uint8_t *)malloc(nDataSize);
  if (!pbData)
    return O2_ERROR_MEM;

  // Извлечение данных потока в буфер
  res = pOLE2File->ExtractStreamData(dwStreamID, nDataOffset, nDataSize,
                                     pbData, &nBytesRead);

  const void *pPackedMacroSrc;
  if (bVBA96)
  {
    // Получение адреса исходного кода макроса в потоке версии VBA96
    pPackedMacroSrc = GetVBA96MacroSourcePtr(pbData, nBytesRead);
  }
  else
  {
    pPackedMacroSrc = pbData;
  }

  size_t cbPackedMacroSrc;

  if (pPackedMacroSrc)
  {
    cbPackedMacroSrc = nBytesRead - ((uint8_t *)pPackedMacroSrc - pbData);
    if (cbPackedMacroSrc >= sizeof(uint32_t))
    {
      // Проверка заголовка запакованного исходного кода макроса
      if ((((uint32_t *)pPackedMacroSrc)[0] & 0x007000FF) != 0x00300001)
      {
        // Поиск запакованного исходного кода макроса
        pPackedMacroSrc = FindVBAMacroSource(pPackedMacroSrc,
                                             cbPackedMacroSrc);
      }
    }
    else
    {
      pPackedMacroSrc = NULL;
    }
  }

  O2Res res2;
  if (pPackedMacroSrc)
  {
    // Распаковка исходного кода макроса VBA
    cbPackedMacroSrc = nBytesRead - ((uint8_t *)pPackedMacroSrc - pbData);
    res2 = DecompressVBAMacroSource(pPackedMacroSrc, cbPackedMacroSrc,
                                    pOutStream, pnProcessedSize);
  }
  else
  {
    res2 = O2_ERROR_DATA;
  }
  if ((res2 != O2_OK) && (res == O2_OK))
    res = res2;

  free(pbData);

  return res;
}
/***************************************************************************/
/* DetectVBAStream - Детектирование потока VBA                             */
/***************************************************************************/
bool DetectVBAStream(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID
  )
{
  if (!pOLE2File || !pOLE2File->GetOpened() ||
      (dwStreamID == NOSTREAM))
    return false;

  const DIRECTORY_ENTRY *pDirEntry;
  pDirEntry = pOLE2File->GetDirEntry(dwStreamID);
  if (!pDirEntry ||
      (pDirEntry->bObjectType != OBJ_TYPE_STREAM))
    return false;

  // Проверка вхождения потока в каталог VBA
  uint32_t dwParentDirEntryID;
  dwParentDirEntryID = pOLE2File->GetParentDirEntryID(dwStreamID);
  if (dwParentDirEntryID == NOSTREAM)
    return false;
  const DIRECTORY_ENTRY *pParentDirEntry;
  pParentDirEntry = pOLE2File->GetDirEntry(dwParentDirEntryID);
  if (!pParentDirEntry ||
      (pParentDirEntry->bObjectType != OBJ_TYPE_STORAGE) ||
      (pParentDirEntry->wSizeOfName != sizeof(VBA_STORAGE_NAME)) ||
       StrCmpNIW((const wchar_t *)pParentDirEntry->bName, VBA_STORAGE_NAME,
                 (sizeof(VBA_STORAGE_NAME) >> 1) - 1))
    return false;

  // Чтение заголовка потока VBA
  size_t nBytesRead;
  VBA_STREAM_HEADER vbaStreamHdr;
  pOLE2File->ExtractStreamData(dwStreamID, 0, sizeof(vbaStreamHdr),
                               &vbaStreamHdr, &nBytesRead);
  if ((nBytesRead < sizeof(vbaStreamHdr.wSign)) ||
      (vbaStreamHdr.wSign != VBA_STREAM_SIGNATURE))
    return false;
  if (nBytesRead <= sizeof_through_field(VBA_STREAM_HEADER, bVersion))
    return false;

  return true;
}
/***************************************************************************/
/* MakeVBAMacroSourceFileName - Получение имени файла исходного кода       */
/*                              макроса VBA                                */
/***************************************************************************/
size_t MakeVBAMacroSourceFileName(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  wchar_t *pBuffer,
  size_t nBufferSize
  )
{
  if (!pOLE2File || !pOLE2File->GetOpened() ||
      (dwStreamID == NOSTREAM))
    return 0;

  const DIRECTORY_ENTRY *pDirEntry;
  pDirEntry = pOLE2File->GetDirEntry(dwStreamID);
  if (!pDirEntry ||
      (pDirEntry->bObjectType != OBJ_TYPE_STREAM))
    return 0;

  const wchar_t *pwchName;
  size_t cchName;

  // Поиск псевдонима имени потока
  pwchName = pOLE2File->FindDirEntryNameAlias(dwStreamID);
  if (pwchName && pwchName[0])
  {
#ifdef USE_WINAPI
    cchName = ::lstrlen(pwchName);
#else
    cchName = wcslen(pwchName);
#endif  // USE_WINAPI
  }
  else
  {
    pwchName = (wchar_t *)pDirEntry->bName;
    cchName = pDirEntry->wSizeOfName >> 1;
    if (cchName > (MAX_DIR_ENTRY_NAME_SIZE >> 1))
      cchName = (MAX_DIR_ENTRY_NAME_SIZE >> 1);
    if (cchName != 0)
      cchName--;
  }

  if (cchName == 0)
  {
    pwchName = VBA_MODULE_FILE_NAME;
    cchName = countof(VBA_MODULE_FILE_NAME) - 1;
  }

  size_t cchExt = countof(VBA_MACRO_FILE_EXT) - 1;
  size_t cchFileName = cchName + cchExt;

  if (!pBuffer || (nBufferSize == 0))
    return (cchFileName + 1);

  if (nBufferSize <= cchFileName)
  {
    if (nBufferSize <= cchExt + 1)
    {
      if (nBufferSize <= cchName)
        cchName = nBufferSize - 1;
      cchExt = nBufferSize - cchName - 1;
    }
    else
    {
      cchName = nBufferSize - cchExt - 1;
    }
  }

  if (cchName != 0)
    memcpy(pBuffer, pwchName, cchName * sizeof(wchar_t));
  if (cchExt != 0)
    memcpy(&pBuffer[cchName], VBA_MACRO_FILE_EXT, cchExt * sizeof(wchar_t));
  pBuffer[cchName + cchExt] = L'\0';

  if (nBufferSize <= cchFileName)
    return (cchFileName + 1);
  return cchFileName;
}
//---------------------------------------------------------------------------
