//---------------------------------------------------------------------------
#define WIN32_LEAN_AND_MEAN  // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#include <malloc.h>
#include <memory.h>
#include "StrUtils.h"
#include "OLE2CF.h"
#include "OLE2Native.h"
//---------------------------------------------------------------------------
#ifndef countof
#define countof(a)  (sizeof(a)/sizeof(a[0]))
#endif  // countof
//---------------------------------------------------------------------------
// ������ ������ ��� ������ ������
#define STREAM_DATA_BUFFER_SIZE  512
//---------------------------------------------------------------------------
// ��� ����� ���������� �������
const wchar_t EMBEDDED_FILE_NAME[] = L"{embedded}";
// ���������� ������ ��� ��������� ����� ������ ���������� �������
const wchar_t PBRUSH_FILE_EXT[]    = L".bmp";
const wchar_t WAVESOUND_FILE_EXT[] = L".wav";

// ������������ ������ "\1CompObj"
const wchar_t COMPOBJ_STREAM_NAME[] = L"\1CompObj";

// �������� ������ � ����� ���������� ������� � ������ "\1CompObj"
#define COMPOBJ_NATIVE_DATA_TYPE_OFFSET  0x1C

// ���� ���������� �������
const char* const NATIVE_DATA_TYPES[] =
{
  "Package",
  "PBrush",
  "SoundRec"
};
//---------------------------------------------------------------------------
/***************************************************************************/
/*                                                                         */
/*   COLE2NativeDataStream - ����� ������ ���������� ������� ����� OLE2    */
/*                           (������ "\1Ole10Native")                      */
/*                                                                         */
/***************************************************************************/

/***************************************************************************/
/* COLE2NativeDataStream - ����������� ������                              */
/***************************************************************************/
COLE2NativeDataStream::COLE2NativeDataStream()
  : m_bOpened(false),
    m_nativeDataType(NATIVE_DATA_UNKNOWN),
    m_pOLE2File(NULL),
    m_dwStreamID(NOSTREAM),
    m_nStreamDataSize(0)
{
  m_nativeData.nOffset = 0;
  m_nativeData.nLength = 0;
  memset(&m_pkgInfo, 0, sizeof(OLEPackageInfo));
}
/***************************************************************************/
/* ~COLE2NativeDataStream - ���������� ������                              */
/***************************************************************************/
COLE2NativeDataStream::~COLE2NativeDataStream()
{
  // ��������
  Close();
}
/***************************************************************************/
/* Open - ��������                                                         */
/***************************************************************************/
O2Res COLE2NativeDataStream::Open(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID
  )
{
  // ��������
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
      (pDirEntry->dwLowStreamSize < sizeof(uint32_t)))
    return O2_ERROR_PARAM;

  O2Res res;

  uint32_t dwNativeDataSize;

  // ��������� ������� ������
  res = pOLE2File->ExtractStreamData(dwStreamID, 0, sizeof(dwNativeDataSize),
                                     &dwNativeDataSize, NULL);
  if ((res != O2_OK) ||
      (dwNativeDataSize == 0))
    return res;

  if (dwNativeDataSize > pDirEntry->dwLowStreamSize - sizeof(uint32_t))
    return O2_ERROR_DATA;

  m_pOLE2File = pOLE2File;
  m_dwStreamID = dwStreamID;
  m_nStreamDataSize = dwNativeDataSize + sizeof(uint32_t);

  // ����������� ���� ������ ���������� �������
  res = DetectNativeDataType();

  if ((res != O2_OK) ||
      (m_nativeDataType == NATIVE_DATA_PACKAGE))
  {
    // �������� (OLE Package)
    res = OpenPackage();
    if (m_bOpened)
    {
      m_nativeDataType = NATIVE_DATA_PACKAGE;
      return res;
    }
    if (m_nativeDataType == NATIVE_DATA_PACKAGE)
    {
      // ��������
      Close();
      return res;
    }
    // ������� ���������� � ��������� ������� (OLE Package)
    memset(&m_pkgInfo, 0, sizeof(OLEPackageInfo));
  }

  m_nativeData.nOffset = sizeof(uint32_t);
  m_nativeData.nLength = dwNativeDataSize;
  m_bOpened = true;
  return O2_OK;
}
/***************************************************************************/
/* OpenPackage - �������� (OLE Package)                                    */
/***************************************************************************/
O2Res COLE2NativeDataStream::OpenPackage()
{
  O2Res res;

  uint8_t buf[STREAM_DATA_BUFFER_SIZE];
  m_pBuffer = buf;
  m_nBufferSize = sizeof(buf);

  // ������ ������ ������ � �����
  res = ReadStreamData(0);
  if (m_nBytesRead < sizeof(uint32_t) + sizeof(uint16_t))
    return (res != O2_OK) ? res : O2_ERROR_DATA;

  // Flags1 (=0x0002)
  uint16_t wFlags1 = ((uint16_t *)(buf + sizeof(uint32_t)))[0];

  m_nCurrentPos = sizeof(uint32_t) + sizeof(uint16_t);

  // ����� ���������� ���������� ������� ������� (ANSI)
  res = ScanFilePathA(&m_pkgInfo.LabelA, &m_pkgInfo.LabelNameA);
  if (res != O2_OK)
    return res;
  // ���� � ����� ���������� ������� (ANSI)
  res = ScanFilePathA(&m_pkgInfo.FilePathA, &m_pkgInfo.FileNameA);
  if (res != O2_OK)
    return res;

  // 8 ������: wFlags2 (0x0000), wUnknown (0x0003), cbTempPath
  res = ScanData(NULL, 8);
  if (res != O2_OK)
    return res;

  // ���� � ���������� ����� ���������� ������� (ANSI)
  res = ScanFilePathA(&m_pkgInfo.TempPathA, &m_pkgInfo.TempNameA);
  if (res != O2_OK)
    return res;

  // ������ ������ ���������� �������
  uint32_t cbNativeData;
  res = ScanData(&cbNativeData, sizeof(cbNativeData));
  if (res != O2_OK)
    return res;
  if (cbNativeData > m_nStreamDataSize - m_nCurrentPos)
    return O2_ERROR_DATA;
  m_nativeData.nOffset = m_nCurrentPos;
  m_nativeData.nLength = cbNativeData;

  m_bOpened = true;

  m_nCurrentPos += cbNativeData;
  if (m_nCurrentPos >= m_nStreamDataSize - sizeof(uint32_t))
    return O2_OK;

  // ������ ������ ������ � �����
  res = ReadStreamData(m_nCurrentPos);
  if (m_nBytesRead < sizeof(uint32_t))
    return (res != O2_OK) ? res : O2_ERROR_DATA;

  // ���� � ���������� ����� ���������� ������� (Unicode)
  res = ScanFilePathW(&m_pkgInfo.TempPathW, &m_pkgInfo.TempNameW);
  if (res != O2_OK)
    return res;
  // ����� ���������� ������� (Unicode)
  res = ScanFilePathW(&m_pkgInfo.LabelW, &m_pkgInfo.LabelNameW);
  if (res != O2_OK)
    return res;
  // ���� � ����� ���������� ������� (Unicode)
  res = ScanFilePathW(&m_pkgInfo.FilePathW, &m_pkgInfo.FileNameW);
  return res;
}
/***************************************************************************/
/* Close - ��������                                                        */
/***************************************************************************/
void COLE2NativeDataStream::Close()
{
  m_bOpened = false;
  m_nativeDataType = NATIVE_DATA_UNKNOWN;
  m_nativeData.nOffset = 0;
  m_nativeData.nLength = 0;
  m_pOLE2File = NULL;
  m_dwStreamID = NOSTREAM;
  m_nStreamDataSize = 0;
  memset(&m_pkgInfo, 0, sizeof(OLEPackageInfo));
}
/***************************************************************************/
/* GetPackageFileName - ��������� ����� ����� ���������� �������           */
/***************************************************************************/
size_t COLE2NativeDataStream::GetPackageFileName(
  wchar_t *pBuffer,
  size_t nBufferSize
  ) const
{
  const O2DataRange *pFileNameRange;
  bool bDoConvert = false;
  if (m_pkgInfo.LabelNameW.nLength != 0)
  {
    pFileNameRange = &m_pkgInfo.LabelNameW;
  }
  else if (m_pkgInfo.LabelNameA.nLength != 0)
  {
    pFileNameRange = &m_pkgInfo.LabelNameA;
    bDoConvert = true;
  }
  else if (m_pkgInfo.FileNameW.nLength != 0)
  {
    pFileNameRange = &m_pkgInfo.FileNameW;
  }
  else if (m_pkgInfo.FileNameA.nLength != 0)
  {
    pFileNameRange = &m_pkgInfo.FileNameA;
    bDoConvert = true;
  }
  else if (m_pkgInfo.TempNameW.nLength != 0)
  {
    pFileNameRange = &m_pkgInfo.TempNameW;
  }
  else if (m_pkgInfo.TempNameA.nLength != 0)
  {
    pFileNameRange = &m_pkgInfo.TempNameA;
    bDoConvert = true;
  }
  else
  {
    return 0;
  }

  if (!bDoConvert)
  {
    // ���������� ������ (Unicode)
    return ExtractStringW(pFileNameRange, pBuffer, nBufferSize);
  }

  char *pszFileName = (char *)malloc((pFileNameRange->nLength + 1) *
                                     sizeof(char));
  if (!pszFileName)
    return 0;

  // ���������� ������ (ANSI)
  ExtractStringA(pFileNameRange, pszFileName, pFileNameRange->nLength + 1);

  // ����������� ������������ ������� ������
  size_t nWideCharCount = ::MultiByteToWideChar(CP_ACP, 0, pszFileName, -1,
                                                NULL, 0);

  if ((nWideCharCount != 0) && pBuffer && (nBufferSize != 0))
  {
    memset(pBuffer, 0, nBufferSize * sizeof(wchar_t));
    if (nBufferSize > 1)
    {
      // �������������� ����� ����� �� ��������� ANSI � Unicode
      int cch = ::MultiByteToWideChar(CP_ACP, 0, pszFileName, -1, pBuffer,
                                      (int)(nBufferSize - 1));
      if (cch != 0) nWideCharCount = cch - 1;
    }
  }

  free(pszFileName);

  return nWideCharCount;
}
/***************************************************************************/
/* ExtractNativeData - ���������� ������ ���������� �������                */
/***************************************************************************/
O2Res COLE2NativeDataStream::ExtractNativeData(
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  ) const
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  if (!m_pOLE2File ||
      (m_dwStreamID == NOSTREAM))
    return O2_ERROR_PARAM;

  if (m_nativeData.nLength == 0)
    return O2_OK;

  O2Res res = O2_OK;

  if (pOutStream)
  {
    // ���������� ����������� ������
    uint64_t nProcessedSize;
    res = m_pOLE2File->ExtractStreamData(m_dwStreamID, m_nativeData.nOffset,
                                         m_nativeData.nLength, pOutStream,
                                         &nProcessedSize);
    if (pnProcessedSize)
      *pnProcessedSize = (size_t)nProcessedSize;
  }
  else
  {
    if (pnProcessedSize)
      *pnProcessedSize = m_nativeData.nLength;
  }

  return res;
}
/***************************************************************************/
/* GetNativeDataTypeString - ��������� ������ � ������������� ���� ������  */
/*                           ���������� �������                            */
/***************************************************************************/
O2Res GetNativeDataTypeString(
  const void *pDataTypeInfo,
  size_t cbDataTypeInfo,
  PO2DataRange pDataTypeStrRange
  )
{
  const uint8_t *p = (uint8_t *)pDataTypeInfo;
  size_t cb = cbDataTypeInfo;

  // 0 - ������������ ���� ���������� �������
  // 1 - ������ ������ ������
  // 2 - ������������� ��������� (ProgID)
  for (size_t i = 0; i < 3; i++)
  {
    if (cb < sizeof(uint32_t))
      return O2_ERROR_DATA;
    uint32_t dwStrLen = ((uint32_t *)p)[0];
    p += sizeof(uint32_t);
    cb -= sizeof(uint32_t);
    if (dwStrLen == 0)
      continue;
    if (i == 1)
    {
      // ������ ������ ������
      if ((dwStrLen == 0xFFFFFFFF) || (dwStrLen == 0xFFFFFFFE))
      {
        // ������������� ������� ������ ������
        if (cb < sizeof(uint32_t))
          return O2_ERROR_DATA;
        uint32_t dwClipboardFormatId = ((uint32_t *)p)[0];
        p += sizeof(uint32_t);
        cb -= sizeof(uint32_t);
        continue;
      }
    }
    else if (i == 2)
    {
      // ������������� ��������� (ProgID)
      // ����� ������ <= 40
      if (dwStrLen > 40)
        break;
    }
    if (cb < dwStrLen)
      return O2_ERROR_DATA;
    if (dwStrLen > 1)
    {
      pDataTypeStrRange->nOffset = p - (uint8_t *)pDataTypeInfo;
      pDataTypeStrRange->nLength = dwStrLen;
    }
    p += dwStrLen;
    cb -= dwStrLen;
  }

  return O2_OK;
}
/***************************************************************************/
/* DetectNativeDataType - ����������� ���� ������ ���������� �������       */
/***************************************************************************/
O2Res COLE2NativeDataStream::DetectNativeDataType()
{
  m_nativeDataType = NATIVE_DATA_UNKNOWN;

  if (!m_pOLE2File ||
      (m_dwStreamID == NOSTREAM))
    return O2_ERROR_PARAM;

  // ����� ������ "\1CompObj" � ��� �� ��������
  uint32_t dwStorageID = m_pOLE2File->GetParentDirEntryID(m_dwStreamID);
  if (dwStorageID == NOSTREAM)
    return O2_ERROR_PARAM;
  uint32_t dwStreamID = m_pOLE2File->FindDirEntry(dwStorageID,
                                                  OBJ_TYPE_STREAM,
                                                  COMPOBJ_STREAM_NAME,
                                                  false);
  if (dwStreamID == NOSTREAM)
    return O2_ERROR_PARAM;

  const DIRECTORY_ENTRY *pDirEntry;
  pDirEntry = m_pOLE2File->GetDirEntry(dwStreamID);
  if (!pDirEntry)
    return O2_ERROR_PARAM;

  uint64_t nStreamSize = (m_pOLE2File->GetMajorVersion() > 3)
                           ? pDirEntry->qwStreamSize
                           : pDirEntry->dwLowStreamSize;
  if ((nStreamSize > V3_MAX_STREAM_SIZE) ||
      (pDirEntry->dwLowStreamSize < COMPOBJ_NATIVE_DATA_TYPE_OFFSET +
                                    sizeof(uint32_t)))
    return O2_ERROR_PARAM;

  O2Res res;

  uint8_t buf[STREAM_DATA_BUFFER_SIZE];
  size_t nBytesRead;

  // ������ ���������� � ���� ���������� �������
  size_t nBytesToRead = pDirEntry->dwLowStreamSize -
                        COMPOBJ_NATIVE_DATA_TYPE_OFFSET;
  if (nBytesToRead > sizeof(buf)) nBytesToRead = sizeof(buf);
  res = m_pOLE2File->ExtractStreamData(dwStreamID,
                                       COMPOBJ_NATIVE_DATA_TYPE_OFFSET,
                                       nBytesToRead, buf, &nBytesRead);

  O2DataRange dataTypeStr;
  dataTypeStr.nLength = 0;

  // ��������� ������ � ������������� ���� ������ ���������� �������
  res = GetNativeDataTypeString(buf, nBytesRead, &dataTypeStr);
  if (dataTypeStr.nLength == 0)
    return (res != O2_OK) ? res : O2_ERROR_DATA;

  char *pszDataType = (char *)buf + dataTypeStr.nOffset;
  pszDataType[dataTypeStr.nLength - 1] = '\0';

  for (size_t i = 0; i < countof(NATIVE_DATA_TYPES); i++)
  {
    if (!StrCmpIA(pszDataType, NATIVE_DATA_TYPES[i]))
    {
      m_nativeDataType = (NativeDataType)(NATIVE_DATA_PACKAGE + i);
      break;
    }
  }

  return O2_OK;
}
/***************************************************************************/
/* ExtractStringA - ���������� ������ (ANSI)                               */
/***************************************************************************/
size_t COLE2NativeDataStream::ExtractStringA(
  const O2DataRange *pStrRange,
  char *pBuffer,
  size_t nBufferSize
  ) const
{
  if (!m_pOLE2File ||
      (m_dwStreamID == NOSTREAM) ||
      (pStrRange->nLength == 0))
    return 0;

  if (!pBuffer || (nBufferSize == 0))
    return (pStrRange->nLength + 1);

  O2Res res = O2_OK;

  size_t cch = pStrRange->nLength;
  if (cch >= nBufferSize)
    cch = nBufferSize - 1;
  if (cch != 0)
  {
    // ���������� ����������� ������
    res = m_pOLE2File->ExtractStreamData(m_dwStreamID, pStrRange->nOffset,
                                         cch, pBuffer, &cch);
  }
  pBuffer[cch] = '\0';

  if (res != O2_OK)
    return 0;
  return (pStrRange->nLength < nBufferSize) ? pStrRange->nLength
                                            : (pStrRange->nLength + 1);
}
/***************************************************************************/
/* ExtractStringW - ���������� ������ (Unicode)                            */
/***************************************************************************/
size_t COLE2NativeDataStream::ExtractStringW(
  const O2DataRange *pStrRange,
  wchar_t *pBuffer,
  size_t nBufferSize
  ) const
{
  if (!m_pOLE2File ||
      (m_dwStreamID == NOSTREAM) ||
      (pStrRange->nLength == 0))
    return 0;

  if (!pBuffer || (nBufferSize == 0))
    return (pStrRange->nLength + 1);

  O2Res res = O2_OK;

  size_t cch = pStrRange->nLength;
  if (cch >= nBufferSize)
    cch = nBufferSize - 1;
  if (cch != 0)
  {
    // ���������� ����������� ������
    size_t cb;
    res = m_pOLE2File->ExtractStreamData(m_dwStreamID, pStrRange->nOffset,
                                         cch << 1, pBuffer, &cb);
    cch = cb >> 1;
  }
  pBuffer[cch] = L'\0';

  if (res != O2_OK)
    return 0;
  return (pStrRange->nLength < nBufferSize) ? pStrRange->nLength
                                            : (pStrRange->nLength + 1);
}
/***************************************************************************/
/* ScanFilePathA - ������������ ���� � ����� (ANSI)                        */
/***************************************************************************/
O2Res COLE2NativeDataStream::ScanFilePathA(
  PO2DataRange pFilePathRange,
  PO2DataRange pFileNameRange
  )
{
  if (m_nCurrentPos < m_nBufferStartPos)
    return O2_ERROR_INTERNAL;
  size_t nBufPos = m_nCurrentPos - m_nBufferStartPos;
  if (nBufPos > m_nBytesRead)
    return O2_ERROR_INTERNAL;
  if (m_nBytesRead == 0)
    return O2_ERROR_DATA;

  O2Res res = O2_OK;

  size_t nPathOfs = m_nCurrentPos;
  size_t nNameOfs = nPathOfs;

  do
  {
    // ����� ����������� ���� � ����� ������
    const uint8_t *pDelim = NULL;
    const uint8_t *pch = m_pBuffer + nBufPos;
    size_t cch = m_nBytesRead - nBufPos;
    while (cch != 0)
    {
      if (*pch == '\0')
        break;
      if ((*pch == '\\') || (*pch == '/') || (*pch == ':'))
        pDelim = pch;
      pch++;
      cch--;
    }
    if (pDelim)
      nNameOfs = m_nBufferStartPos + (pDelim + 1 - m_pBuffer);
    if (cch != 0)
    {
      m_nCurrentPos = m_nBufferStartPos + (pch - m_pBuffer);
      res = O2_OK;
      break;
    }
    if (res != O2_OK)
      break;
    // ������ ��������� ������ ������ � �����
    res = ReadNextStreamData();
    nBufPos = 0;
  }
  while (m_nBytesRead != 0);

  if (res != O2_OK)
    return res;
  if (m_nBytesRead == 0)
    return O2_ERROR_DATA;

  pFilePathRange->nOffset = nPathOfs;
  pFilePathRange->nLength = m_nCurrentPos - nPathOfs;
  pFileNameRange->nOffset = nNameOfs;
  pFileNameRange->nLength = m_nCurrentPos - nNameOfs;
  m_nCurrentPos++;
  return O2_OK;
}
/***************************************************************************/
/* ScanFilePathW - ������������ ���� � ����� (Unicode)                     */
/***************************************************************************/
O2Res COLE2NativeDataStream::ScanFilePathW(
  PO2DataRange pFilePathRange,
  PO2DataRange pFileNameRange
  )
{
  O2Res res;
  uint32_t dwPathLen;

  // ��������� ����� ������
  res = ScanData(&dwPathLen, sizeof(dwPathLen));
  if (res != O2_OK)
    return res;

  size_t nBufPos = m_nCurrentPos - m_nBufferStartPos;
  // �������� �������� � ������?
  if (nBufPos & 1)
    return O2_ERROR_INTERNAL;

  size_t nPathOfs = m_nCurrentPos;
  size_t nNameOfs = nPathOfs;

  size_t cchPath = dwPathLen;
  do
  {
    // ����� ����������� ���� � ����� ������
    const uint16_t *pDelim = NULL;
    const uint16_t *pwch;
    pwch = (uint16_t *)(m_pBuffer + nBufPos);
    size_t cch = (m_nBytesRead - nBufPos) >> 1;
    if (cch > cchPath) cch = cchPath;
    cchPath -= cch;
    while (cch != 0)
    {
      if ((*pwch == L'\\') || (*pwch == L'/') || (*pwch == L':'))
        pDelim = pwch;
      pwch++;
      cch--;
    }
    if (pDelim)
      nNameOfs = m_nBufferStartPos + ((uint8_t *)(pDelim + 1) - m_pBuffer);
    if (cchPath == 0)
    {
      res = O2_OK;
      break;
    }
    if (res != O2_OK)
      break;
    // ������ ��������� ������ ������ � �����
    res = ReadNextStreamData();
    nBufPos = 0;
  }
  while (m_nBytesRead != 0);

  if (res != O2_OK)
    return res;
  if (m_nBytesRead == 0)
    return O2_ERROR_DATA;

  pFilePathRange->nOffset = nPathOfs;
  pFilePathRange->nLength = dwPathLen;
  pFileNameRange->nOffset = nNameOfs;
  pFileNameRange->nLength = dwPathLen - ((nNameOfs - nPathOfs) >> 1);
  m_nCurrentPos += (dwPathLen << 1);
  return O2_OK;
}
/***************************************************************************/
/* ScanData - ������������ ������                                          */
/***************************************************************************/
O2Res COLE2NativeDataStream::ScanData(
  void *buf,
  size_t count
  )
{
  if (m_nCurrentPos < m_nBufferStartPos)
    return O2_ERROR_INTERNAL;
  size_t nBufPos = m_nCurrentPos - m_nBufferStartPos;
  if (nBufPos > m_nBytesRead)
    return O2_ERROR_INTERNAL;

  if (count == 0)
    return O2_OK;

  O2Res res = O2_OK;

  while (m_nBytesRead != 0)
  {
    size_t cbPart = count;
    if (cbPart > m_nBytesRead - nBufPos)
      cbPart = m_nBytesRead - nBufPos;
    if (buf && (cbPart != 0))
    {
      memcpy(buf, m_pBuffer + nBufPos, cbPart);
      buf = (uint8_t *)buf + cbPart;
    }
    m_nCurrentPos += cbPart;
    count -= cbPart;
    if (count == 0)
      return O2_OK;
    if (res != O2_OK)
      break;
    // ������ ��������� ������ ������ � �����
    res = ReadNextStreamData();
    nBufPos = 0;
  }

  return (res != O2_OK) ? res : O2_ERROR_DATA;
}
/***************************************************************************/
/* ReadNextStreamData - ������ ��������� ������ ������ � �����             */
/***************************************************************************/
O2Res COLE2NativeDataStream::ReadNextStreamData()
{
  // ������ ������ ������ � �����
  return ReadStreamData(m_nBufferStartPos + m_nBytesRead);
}
/***************************************************************************/
/* ReadStreamData - ������ ������ ������ � �����                           */
/***************************************************************************/
O2Res COLE2NativeDataStream::ReadStreamData(
  size_t nStartPos
  )
{
  m_nBufferStartPos = nStartPos;
  if (m_nBufferStartPos >= m_nStreamDataSize)
  {
    // EOF
    m_nBytesRead = 0;
    return O2_OK;
  }

  // ���������� ����������� ������
  size_t nBytesToRead = m_nBufferSize;
  if (nBytesToRead > m_nStreamDataSize - m_nBufferStartPos)
    nBytesToRead = m_nStreamDataSize - m_nBufferStartPos;
  return m_pOLE2File->ExtractStreamData(m_dwStreamID, m_nBufferStartPos,
                                        nBytesToRead, m_pBuffer,
                                        &m_nBytesRead);
}
//---------------------------------------------------------------------------
/***************************************************************************/
/* GetEmbeddedObjectInfo - ��������� ���������� � ��������� �������        */
/*                         �� ������ "\1Ole10Native"                       */
/***************************************************************************/
O2Res GetEmbeddedObjectInfo(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  wchar_t *pwszFileNameBuf,
  size_t cchFileNameBuf,
  size_t *pnFileSize
  )
{
  O2Res res;

  // �������� ������� ������ ���������� ������� ����� OLE2
  COLE2NativeDataStream nativeDataStream;

  // �������� ������ ���������� ������� ����� OLE2
  res = nativeDataStream.Open(pOLE2File, dwStreamID);
  if (nativeDataStream.GetOpened())
  {
    // ��������� ������� ������ ���������� �������
    *pnFileSize = nativeDataStream.GetNativeDataSize();
    if (pwszFileNameBuf && (cchFileNameBuf != 0))
    {
      // ��������� ����� ����� ���������� �������
      size_t nRet = nativeDataStream.GetPackageFileName(pwszFileNameBuf,
                                                        cchFileNameBuf);
      if (nRet == 0)
      {
        // ����������� ���������� ����� �� ���� ���������� �������
        const wchar_t *pwszExt;
        NativeDataType nativeDataType;
        nativeDataType = nativeDataStream.GetNativeDataType();
        if (nativeDataType == NATIVE_DATA_PBRUSH)
          pwszExt = PBRUSH_FILE_EXT;
        else if (nativeDataType == NATIVE_DATA_WAVESOUND)
          pwszExt = WAVESOUND_FILE_EXT;
        else
          pwszExt = NULL;
        // ����������� ����� � ���������� �����
        StrCat(EMBEDDED_FILE_NAME, pwszExt, pwszFileNameBuf, cchFileNameBuf);
      }
    }
  }

  return res;
}
/***************************************************************************/
/* ExtractEmbeddedObject - ���������� ���������� ������� �� ������         */
/*                         "\1Ole10Native"                                 */
/***************************************************************************/
O2Res ExtractEmbeddedObject(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  )
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  O2Res res;

  // �������� ������� ������ ���������� ������� ����� OLE2
  COLE2NativeDataStream nativeDataStream;

  // �������� ������ ���������� ������� ����� OLE2
  res = nativeDataStream.Open(pOLE2File, dwStreamID);
  if (nativeDataStream.GetOpened())
  {
    // ���������� ������ ���������� �������
    O2Res res2 = nativeDataStream.ExtractNativeData(pOutStream,
                                                    pnProcessedSize);
    if ((res2 != O2_OK) && (res == O2_OK))
      res = res2;
  }

  return res;
}
//---------------------------------------------------------------------------
