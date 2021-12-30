//---------------------------------------------------------------------------
#define WIN32_LEAN_AND_MEAN  // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#include "OLE2CF.h"
#include "OLE2Word.h"
#include <malloc.h>
#include <memory.h>
//---------------------------------------------------------------------------
// Наименование потока Table
const wchar_t TABLE_STREAM_NAME[] = L"0Table";
//---------------------------------------------------------------------------
/***************************************************************************/
/* WriteTextStream - Запись текста Unicode в поток                         */
/***************************************************************************/
O2Res WriteTextStream(
  const wchar_t *pwchText,
  size_t cchTextLength,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  )
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  if (!pwchText || (cchTextLength == 0))
    return O2_OK;

  size_t cbLength;

  // Определение размера буфера для преобразования текста из UTF-16 в UTF-8
  cbLength = ::WideCharToMultiByte(CP_UTF8, 0, pwchText, (int)cchTextLength,
                                   NULL, 0, NULL, NULL);
  if (cbLength == 0)
    return O2_ERROR_DATA;

  if (!pOutStream)
  {
    if (pnProcessedSize)
      *pnProcessedSize = cbLength;
    return O2_OK;
  }

  // Выделение буфера для преобразования текста из UTF-16 в UTF-8
  char *pBuffer = (char *)malloc(cbLength * sizeof(char));
  if (!pBuffer)
    return O2_ERROR_MEM;

  O2Res res = O2_OK;

  // Преобразование текста из UTF-16 в UTF-8
  cbLength = ::WideCharToMultiByte(CP_UTF8, 0, pwchText, (int)cchTextLength,
                                   pBuffer, (int)cbLength, NULL, NULL);
  if (cbLength == 0)
  {
    res = O2_ERROR_DATA;
  }
  else
  {
    size_t nBytesWritten = pOutStream->Write(pBuffer, cbLength);
    if (pnProcessedSize)
      *pnProcessedSize = nBytesWritten;
    if (nBytesWritten != cbLength)
      res = O2_ERROR_WRITE;
  }

  free(pBuffer);

  return res;
}
/***************************************************************************/
/* WriteTextStream - Запись текста ANSI в поток                            */
/***************************************************************************/
O2Res WriteTextStream(
  const char *pchText,
  size_t cchTextLength,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  )
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  if (!pchText || (cchTextLength == 0))
    return O2_OK;

  size_t cchLength;

  // Определение размера буфера для преобразования текста из ANSI в UTF-16
  cchLength = ::MultiByteToWideChar(CP_ACP, 0, pchText, (int)cchTextLength,
                                    NULL, 0);
  if (cchLength == 0)
    return O2_ERROR_DATA;

  // Выделение буфера для преобразования текста из ANSI в UTF-16
  wchar_t *pBuffer = (wchar_t *)malloc(cchLength * sizeof(wchar_t));
  if (!pBuffer)
    return O2_ERROR_MEM;

  O2Res res = O2_OK;

  // Преобразование текста из ANSI в UTF-16
  cchLength = ::MultiByteToWideChar(CP_ACP, 0, pchText, (int)cchTextLength,
                                    pBuffer, (int)cchLength);
  if (cchLength == 0)
  {
    res = O2_ERROR_DATA;
  }
  else
  {
    // Запись текста Unicode в поток
    res = WriteTextStream(pBuffer, cchLength, pOutStream, pnProcessedSize);
  }

  free(pBuffer);

  return res;
}
//---------------------------------------------------------------------------
/***************************************************************************/
/*                                                                         */
/*          COLE2WordDocumentStream - класс потока документа Word          */
/*                                    (потока "WordDocument")              */
/*                                                                         */
/***************************************************************************/

/***************************************************************************/
/* COLE2WordDocumentStream - Конструктор класса                            */
/***************************************************************************/
COLE2WordDocumentStream::COLE2WordDocumentStream()
  : m_bOpened(false),
    m_pOLE2File(NULL),
    m_dwStreamID(NOSTREAM),
    m_nStreamSize(0),
    m_bWord97(false),
    m_fcText(0),
    m_lcbText(0),
    m_pTextPieces(NULL),
    m_nNumTextPieces(0)
{
}
/***************************************************************************/
/* ~COLE2WordDocumentStream - Деструктор класса                            */
/***************************************************************************/
COLE2WordDocumentStream::~COLE2WordDocumentStream()
{
  // Закрытие
  Close();
}
/***************************************************************************/
/* Open - Открытие                                                         */
/***************************************************************************/
O2Res COLE2WordDocumentStream::Open(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID
  )
{
  // Закрытие
  Close();

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
      (pDirEntry->dwLowStreamSize <= sizeof(FIB)))
    return O2_ERROR_PARAM;

  m_pOLE2File = pOLE2File;
  m_dwStreamID = dwStreamID;
  m_nStreamSize = pDirEntry->dwLowStreamSize;

  // Инициализация
  O2Res res = Init();

  if (!m_bOpened)
  {
    // Закрытие
    Close();
  }

  return res;
}
/***************************************************************************/
/* Init - Инициализация                                                    */
/***************************************************************************/
O2Res COLE2WordDocumentStream::Init()
{
  O2Res res;

  FIB fib;
  size_t nBytesRead;

  // Чтение базовой части заголовка (FibBase) и csw
  res = m_pOLE2File->ExtractStreamData(m_dwStreamID, 0,
                                       sizeof_through_field(FIB, csw),
                                       &fib,
                                       &nBytesRead);
  if ((nBytesRead < sizeof(fib.base.wIdent)) ||
      ((fib.base.wIdent != WORD97_FILE_IDENT) &&
       (fib.base.wIdent != WORD6_FILE_IDENT)))
    return (res != O2_OK) ? res : O2_ERROR_SIGN;
  if (nBytesRead != sizeof_through_field(FIB, csw))
    return (res != O2_OK) ? res : O2_ERROR_DATA;

  // Документ зашифрован?
  if (fib.base.wFlags1 & FIB_FLAG_ENCRYPTED)
    return res;

  m_bWord97 = (fib.base.wIdent == WORD97_FILE_IDENT);

  // Положение текста в потоке (Word6 / Word95 / Word97)
  m_fcText = 0;
  m_lcbText = 0;
  if ((fib.base.fcMin >= sizeof(fib)) &&
      (fib.base.fcMin < m_nStreamSize) &&
      (fib.base.fcMin < fib.base.fcMax))
  {
    size_t fcTextEnd = fib.base.fcMax;
    if (fcTextEnd > m_nStreamSize)
      fcTextEnd = m_nStreamSize;
    m_fcText = fib.base.fcMin;
    m_lcbText = fcTextEnd - m_fcText;
  }

  if (m_bWord97)
  {
    // Word97 и выше
    // Чтение cslw и третьей части заголовка (FibRgLW97)
    res = m_pOLE2File->ExtractStreamData(m_dwStreamID,
                                         sizeof_through_field(FIB, csw) +
                                         (fib.csw << 1),
                                         sizeof(fib.cslw) +
                                         sizeof(fib.fibRgLw),
                                         &fib.cslw,
                                         &nBytesRead);
    if ((nBytesRead != sizeof(fib.cslw) + sizeof(fib.fibRgLw)) ||
        (fib.cslw < (sizeof(fib.fibRgLw) >> 2)))
      return (res != O2_OK) ? res : O2_ERROR_DATA;

    // Поиск текста комплексного документа
    O2Res res2 = ComplexRetrieveText(&fib);
    if ((res2 != O2_OK) && (res == O2_OK))
      res = res2;
  }

  if ((m_nNumTextPieces != 0) ||
      (m_fcText != 0))
    m_bOpened = true;

  return res;
}
/***************************************************************************/
/* ComplexRetrieveText - Поиск текста комплексного документа               */
/***************************************************************************/
O2Res COLE2WordDocumentStream::ComplexRetrieveText(
  const FIB *pFib
  )
{
  O2Res res;

  FIBCLXFCLCB clxFcLcb;
  size_t nBytesRead;

  // Определение смещения полей смещения и размера таблицы Clx
  size_t fcClxFcLcb = sizeof(pFib->base) + sizeof(pFib->csw) +
                      (pFib->csw << 1) + sizeof(pFib->cslw) +
                      (pFib->cslw << 2) + sizeof(uint16_t) +
                      (FIBCLXFCLCB_INDEX << 3);
  if ((m_nStreamSize <= fcClxFcLcb) ||
      (m_nStreamSize - fcClxFcLcb < sizeof(FIBCLXFCLCB)))
    return O2_ERROR_DATA;

  // Получение в потоке документа смещения и размера таблицы Clx
  res = m_pOLE2File->ExtractStreamData(m_dwStreamID, fcClxFcLcb,
                                       sizeof(clxFcLcb), &clxFcLcb,
                                       &nBytesRead);
  if (nBytesRead != sizeof(clxFcLcb))
    return (res != O2_OK) ? res : O2_ERROR_DATA;
  if (clxFcLcb.lcbClx == 0)
    return res;
  if (clxFcLcb.lcbClx < sizeof(uint8_t) + sizeof(uint32_t))
    return (res != O2_OK) ? res : O2_ERROR_DATA;

  // Определение имени потока Table
  wchar_t wszTableStreamName[sizeof(TABLE_STREAM_NAME) / sizeof(wchar_t)];
  memcpy(wszTableStreamName, TABLE_STREAM_NAME, sizeof(TABLE_STREAM_NAME));
  if (pFib->base.wFlags1 & FIB_FLAG_WHICHTBLSTM)
    wszTableStreamName[0]++;

  // Поиск потока Table в том же каталоге
  uint32_t dwStorageID = m_pOLE2File->GetParentDirEntryID(m_dwStreamID);
  if (dwStorageID == NOSTREAM)
    return O2_ERROR_DATA;
  uint32_t dwTableStreamID = m_pOLE2File->FindDirEntry(dwStorageID,
                                                       OBJ_TYPE_STREAM,
                                                       wszTableStreamName,
                                                       false);
  if (dwTableStreamID == NOSTREAM)
    return O2_ERROR_DATA;

  // Проверка положения таблицы Clx в потоке Table
  const DIRECTORY_ENTRY *pDirEntry;
  pDirEntry = m_pOLE2File->GetDirEntry(dwTableStreamID);
  if (!pDirEntry)
    return O2_ERROR_DATA;
  uint64_t nStreamSize = (m_pOLE2File->GetMajorVersion() > 3)
                           ? pDirEntry->qwStreamSize
                           : pDirEntry->dwLowStreamSize;
  if ((nStreamSize > V3_MAX_STREAM_SIZE) ||
      (clxFcLcb.fcClx >= pDirEntry->dwLowStreamSize) ||
      (clxFcLcb.lcbClx > pDirEntry->dwLowStreamSize - clxFcLcb.fcClx))
    return O2_ERROR_DATA;

  uint8_t *pClx = (uint8_t *)malloc(clxFcLcb.lcbClx);
  if (!pClx)
    return O2_ERROR_MEM;

  // Чтение таблицы Clx из потока Table
  res = m_pOLE2File->ExtractStreamData(dwTableStreamID, clxFcLcb.fcClx,
                                       clxFcLcb.lcbClx, pClx, &nBytesRead);
  if (nBytesRead != clxFcLcb.lcbClx)
  {
    if (res != O2_OK) res = O2_ERROR_DATA;
  }
  else
  {
    // Поиск текста комплексного документа
    O2Res res2 = ComplexRetrieveText(pFib, pClx, nBytesRead);
    if ((res2 != O2_OK) && (res == O2_OK))
      res = res2;
  }

  free(pClx);

  return res;
}
/***************************************************************************/
/* ComplexRetrieveText - Поиск текста комплексного документа               */
/***************************************************************************/
O2Res COLE2WordDocumentStream::ComplexRetrieveText(
  const FIB *pFib,
  const void *pClx,
  size_t cbClx
  )
{
  if (cbClx < sizeof(uint8_t))
    return O2_ERROR_DATA;

  const uint8_t *p = (uint8_t *)pClx;
  size_t cb = cbClx;

  if (p[0] == 0x01)
  {
    // Таблица RgPrc
    if (cb < sizeof(uint8_t) + sizeof(uint16_t))
      return O2_ERROR_DATA;
    p++;
    cb -= sizeof(uint8_t) + sizeof(uint16_t);
    size_t cbGrpprl = ((uint16_t *)p)[0];
    if (cbGrpprl > cb)
      return O2_ERROR_DATA;
    p += sizeof(uint16_t) + cbGrpprl;
    cb -= cbGrpprl;
  }

  if ((cb < sizeof(uint8_t) + sizeof(uint32_t)) ||
      (p[0] != 0x02))
    return O2_ERROR_DATA;

  // Таблица Pcdt
  p++;
  cb -= sizeof(uint8_t) + sizeof(uint32_t);
  size_t cbPlcPcd = ((uint32_t *)p)[0];
  if (cbPlcPcd != cb)
    return O2_ERROR_DATA;

  // Определение значения последнего CP
  uint32_t lastCP = pFib->fibRgLw.ccpFtn + pFib->fibRgLw.ccpHdd +
                    pFib->fibRgLw.ccpMcr + pFib->fibRgLw.ccpAtn +
                    pFib->fibRgLw.ccpEdn + pFib->fibRgLw.ccpTxbx +
                    pFib->fibRgLw.ccpHdrTxbx;
  if (lastCP != 0)
    lastCP++;
  lastCP += pFib->fibRgLw.ccpText;

  const uint32_t *pCP = (uint32_t *)(p + sizeof(uint32_t));

  // Определение количества элементов массивов CP и Pcd
  size_t nPieceCount = 0;
  for (; nPieceCount < cbPlcPcd / sizeof(uint32_t); nPieceCount++)
  {
    if (((nPieceCount != 0) &&
         (pCP[nPieceCount] <= pCP[nPieceCount - 1])) ||
        (pCP[nPieceCount] > lastCP))
      return O2_ERROR_DATA;
    if (pCP[nPieceCount] == lastCP)
      break;
  }
  if ((nPieceCount == 0) ||
      (nPieceCount >= cbPlcPcd / sizeof(uint32_t)) ||
      (cbPlcPcd < nPieceCount * (sizeof(uint32_t) + sizeof(PCD)) +
                  sizeof(uint32_t)))
    return O2_ERROR_DATA;

  const PCD *pPcd = (PCD *)(pCP + nPieceCount + 1);

  // Выделение памяти для массива частей текста
  m_pTextPieces = (PWDTextPiece)malloc(nPieceCount * sizeof(WDTextPiece));
  if (!m_pTextPieces)
    return O2_ERROR_MEM;

  O2Res res = O2_OK;

  // Заполнение массива частей текста
  size_t nNumTextPieces = 0;
  for (size_t i = 0; i < nPieceCount; i++)
  {
    size_t fc = pPcd[i].fc & 0x3FFFFFFF;
    size_t lcb = pCP[i + 1] - pCP[i];
    bool bANSI = false;
    if (pPcd[i].fc & PCD_FC_COMPRESSED)
    {
      bANSI = true;
      fc >>= 1;
    }
    else
    {
      lcb <<= 1;
    }
    if (fc >= m_nStreamSize)
    {
      res = O2_ERROR_DATA;
      continue;
    }
    if (lcb > m_nStreamSize - fc)
    {
      res = O2_ERROR_DATA;
      lcb = m_nStreamSize - fc;
      if (!bANSI) lcb &= 1;
      if (lcb == 0)
        continue;
    }
    m_pTextPieces[nNumTextPieces].bANSI = bANSI;
    m_pTextPieces[nNumTextPieces].fc = fc;
    m_pTextPieces[nNumTextPieces].lcb = lcb;
    nNumTextPieces++;
  }
  m_nNumTextPieces = nNumTextPieces;

  return res;
}
/***************************************************************************/
/* Close - Закрытие                                                        */
/***************************************************************************/
void COLE2WordDocumentStream::Close()
{
  m_bOpened = false;
  m_pOLE2File = NULL;
  m_dwStreamID = NOSTREAM;
  m_nStreamSize = 0;
  m_bWord97 = false;
  m_fcText = 0;
  m_lcbText = 0;
  m_nNumTextPieces = 0;
  if (m_pTextPieces)
  {
    // Освобождение выделенной памяти для массива частей текста документа
    free(m_pTextPieces);
    m_pTextPieces = NULL;
  }
}
/***************************************************************************/
/* ExtractText - Извлечение текста                                         */
/***************************************************************************/
O2Res COLE2WordDocumentStream::ExtractText(
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  ) const
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  if (!m_pOLE2File ||
      (m_dwStreamID == NOSTREAM))
    return O2_ERROR_PARAM;

  if (m_pTextPieces && (m_nNumTextPieces != 0))
  {
    // Извлечение текста комплексного документа
    return ComplexExtractText(pOutStream, pnProcessedSize);
  }

  // Извлечение текста из потока (Word6 / Word95 / Word97)
  return W6ExtractText(pOutStream, pnProcessedSize);
}
/***************************************************************************/
/* ComplexExtractText - Извлечение текста комплексного документа           */
/***************************************************************************/
O2Res COLE2WordDocumentStream::ComplexExtractText(
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  ) const
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  if (!m_pTextPieces || (m_nNumTextPieces == 0))
    return O2_ERROR_DATA;

  // Определение размера буфера
  size_t nBufferSize = 0;
  for (size_t i = 0; i < m_nNumTextPieces; i++)
  {
    if (m_pTextPieces[i].lcb > nBufferSize)
      nBufferSize = m_pTextPieces[i].lcb;
  }
  if (nBufferSize == 0)
    return O2_ERROR_DATA;

  // Выделение буфера для частей текста
  void *pBuffer = malloc(nBufferSize);
  if (!pBuffer)
    return O2_ERROR_MEM;

  O2Res res = O2_OK;

  for (size_t i = 0; i < m_nNumTextPieces; i++)
  {
    O2Res res2;
    size_t nBytesRead;
    size_t nProcessedSize;
    // Чтение части текста в буфер
    res2 = m_pOLE2File->ExtractStreamData(m_dwStreamID,
                                          (uint64_t)m_pTextPieces[i].fc,
                                          m_pTextPieces[i].lcb,
                                          pBuffer, &nBytesRead);
    if ((res2 != O2_OK) && (res == O2_OK))
      res = res2;

    if (m_pTextPieces[i].bANSI)
    {
      // Запись текста ANSI в поток
      res2 = WriteTextStream((char *)pBuffer, nBytesRead, pOutStream,
                             &nProcessedSize);
    }
    else
    {
      // Запись текста Unicode в поток
      res2 = WriteTextStream((wchar_t *)pBuffer, nBytesRead >> 1, pOutStream,
                             &nProcessedSize);
    }
    if ((res2 != O2_OK) && (res == O2_OK))
      res = res2;
    if (pnProcessedSize)
      *pnProcessedSize += nProcessedSize;
  }

  free(pBuffer);

  return res;
}
/***************************************************************************/
/* W6ExtractText - Извлечение текста (Word6)                               */
/***************************************************************************/
O2Res COLE2WordDocumentStream::W6ExtractText(
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  ) const
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  if (m_fcText == 0)
    return O2_ERROR_DATA;
  if (m_lcbText == 0)
    return O2_OK;

  O2Res res = O2_OK;

  if (!m_bWord97)
  {
    // Выделение буфера для текста
    void *pBuffer = malloc(m_lcbText);
    if (!pBuffer)
      return O2_ERROR_MEM;

    size_t nBytesRead;

    // Чтение части текста в буфер
    res = m_pOLE2File->ExtractStreamData(m_dwStreamID, (uint64_t)m_fcText,
                                         m_lcbText, pBuffer, &nBytesRead);
    // Запись текста ANSI в поток
    O2Res res2 = WriteTextStream((char *)pBuffer, nBytesRead, pOutStream,
                                 pnProcessedSize);
    if ((res2 != O2_OK) && (res == O2_OK))
      res = res2;

    free(pBuffer);
  }
  else
  {
    if (pOutStream)
    {
      // Извлечение содержимого потока
      uint64_t nProcessedSize;
      res = m_pOLE2File->ExtractStreamData(m_dwStreamID,
                                           (uint64_t)m_fcText,
                                           (uint64_t)m_lcbText,
                                           pOutStream,
                                           &nProcessedSize);
      if (pnProcessedSize)
        *pnProcessedSize = (size_t)nProcessedSize;
    }
    else
    {
      if (pnProcessedSize)
        *pnProcessedSize = m_lcbText;
    }
  }

  return res;
}
//---------------------------------------------------------------------------
/***************************************************************************/
/* ExtractWordDocumentText - Извлечение текста из потока документа Word    */
/***************************************************************************/
O2Res ExtractWordDocumentText(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  )
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  O2Res res;

  // Создание объекта потока документа Word
  COLE2WordDocumentStream wdStream;

  // Открытие потока документа Word
  res = wdStream.Open(pOLE2File, dwStreamID);
  if (wdStream.GetOpened())
  {
    // Извлечение текста
    O2Res res2 = wdStream.ExtractText(pOutStream, pnProcessedSize);
    if ((res2 != O2_OK) && (res == O2_OK))
      res = res2;
  }

  return res;
}
//---------------------------------------------------------------------------
