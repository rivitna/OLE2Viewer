//---------------------------------------------------------------------------
#ifdef USE_WINAPI
#define WIN32_LEAN_AND_MEAN  // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#else
#include <string.h>
#endif  // USE_WINAPI

#include <malloc.h>
#include <memory.h>
#include "OLE2File.h"
//---------------------------------------------------------------------------
#ifndef min
#define min(a, b)  (((a) < (b)) ? (a) : (b))
#endif
//---------------------------------------------------------------------------
// ����������� ����
#define PATH_DELIMITER      L'\\'
#define ISPATHDELIMITER(c)  ((c == L'\\') || (c == L'/'))

// ������������ ���������� �������� ��� "����������" �������
// sizeof(LOST_DIR_NAME) <= MAX_DIR_ENTRY_NAME_SIZE
const wchar_t LOST_DIR_NAME[] = L"!LOST.DIR";
//---------------------------------------------------------------------------
/***************************************************************************/
/* ReadStream - ������ ������ ������                                       */
/***************************************************************************/
O2Res ReadStream(
  CInStream *pInStream,
  uint64_t nOffset,
  size_t nBytesToRead,
  void *pBuffer,
  size_t *pnBytesRead
  )
{
  if (pnBytesRead)
    *pnBytesRead = 0;

  // ��������� ��������� ������
  if (pInStream->Seek(&nOffset, STM_SEEK_SET))
    return O2_ERROR_DATA;

  if (nBytesToRead == 0)
    return O2_OK;

  O2Res res = O2_OK;

  // ������ ������ ������
  size_t nSize = nBytesToRead;
  if (pInStream->Read(pBuffer, &nSize))
  {
    res = O2_ERROR_READ;
  }
  else
  {
    if (nSize != nBytesToRead)
      res = O2_ERROR_DATA;
  }
  if (pnBytesRead)
    *pnBytesRead = nSize;

  return res;
}
//---------------------------------------------------------------------------
/***************************************************************************/
/*                                                                         */
/*                      COLE2File - ����� ����� OLE2                       */
/*                                                                         */
/***************************************************************************/

/***************************************************************************/
/* COLE2File - ����������� ������                                          */
/***************************************************************************/
COLE2File::COLE2File()
  : m_bOpened(false),
    m_pInStream(NULL),
    m_nMajorVersion(0),
    m_nMinorVersion(0),
    m_nEffectiveNumSectors(0),
    m_nEffectiveNumMiniSectors(0),
    m_dwRootEntryStartSector(ENDOFCHAIN),
    m_nSectorSize(0),
    m_nMiniSectorSize(0),
    m_nMiniStreamCutoffSize(0),
    m_pFAT(NULL),
    m_nNumFATEntries(0),
    m_pMiniFAT(NULL),
    m_nNumMiniFATEntries(0),
    m_pDirEntries(NULL),
    m_nNumDirEntries(0),
    m_bSomeDirEntriesLost(false),
    m_nFileSize(0i64),
    m_pDirEntryItems(NULL),
    m_nNumDirEntryItems(0),
    m_pLostStorageDirEntry(NULL),
    m_pDirEntryNameAliases(NULL),
    m_nNumDirEntryNameAliases(0)
{
  memset(&m_fileHeader, 0, COMPOUND_FILE_HEADER_SIZE);
}
/***************************************************************************/
/* ~COLE2File - ���������� ������                                          */
/***************************************************************************/
COLE2File::~COLE2File()
{
  // ��������
  Close();
}
/***************************************************************************/
/* Open - ��������                                                         */
/***************************************************************************/
O2Res COLE2File::Open(
  CInStream *pInStream
  )
{
  // ��������
  Close();

  // �������������
  O2Res res = Init(pInStream);

  if (!m_bOpened)
  {
    // ��������
    Close();
  }

  return res;
}
/***************************************************************************/
/* Init - �������������                                                    */
/***************************************************************************/
O2Res COLE2File::Init(
  CInStream *pInStream
  )
{
  if (!pInStream)
    return O2_ERROR_PARAM;

  uint64_t nFileSize = 0i64;

  // ����������� ������� �����
  if (pInStream->Seek(&nFileSize, STM_SEEK_END) ||
      (nFileSize < COMPOUND_FILE_HEADER_SIZE))
    return O2_ERROR_DATA;

  O2Res res;

  // ������ ��������� ���������� ����� OLE2
  res = ReadStream(pInStream, 0i64, COMPOUND_FILE_HEADER_SIZE, &m_fileHeader,
                   NULL);
  if (res != O2_OK)
    return res;

  // �������� ��������� ���������� ����� OLE2
  if (!CheckFileHeader())
    return O2_ERROR_DATA;

  m_pInStream = pInStream;
  m_nFileSize = nFileSize;

  // ��������� ������
  m_nMajorVersion = m_fileHeader.wMajorVersion;
  m_nMinorVersion = m_fileHeader.wMinorVersion;

  // ��������� �������� ��������
  m_nSectorSize = (size_t)1 << m_fileHeader.wSectorShift;
  m_nMiniSectorSize = (size_t)1 << m_fileHeader.wMiniSectorShift;
  // �������� ���� dwMiniStreamCutoffSize ������������
  m_nMiniStreamCutoffSize = 0x1000;  // m_fileHeader.dwMiniStreamCutoffSize;

  // ��������� ������������ ���������� �������� � �����
  m_nEffectiveNumSectors = (size_t)(m_nFileSize / m_nSectorSize);
  if (m_nEffectiveNumSectors <= 1)
    return O2_ERROR_DATA;
  m_nEffectiveNumSectors--;

  // ��������� ������ � ������ FAT
  res = AllocAndReadFAT();
  if ((res != O2_OK) && (m_nNumFATEntries == 0))
    return res;

  O2Res res2;

  // ��������� ������ � ������ mini FAT
  res2 = AllocAndReadMiniFAT();
  if ((res2 != O2_OK) && (res == O2_OK))
    res = res2;

  // ��������� ������ � ������ ��������
  res2 = AllocAndReadDirectory();
  if (res2 != O2_OK)
  {
    if ((m_nNumDirEntries == 0) && (m_nNumDirEntries >= MAXREGSID))
      return res2;
    if (res == O2_OK)
      res = res2;
  }

  // ��������� ������ ���������� ������� �������� ������
  m_dwRootEntryStartSector = m_pDirEntries[0].dwStartSector;
  // ��������� ������������ ���������� �������� � mini FAT
  uint64_t nRootEntrySize = (m_nMajorVersion > 3)
                              ? m_pDirEntries[0].qwStreamSize
                              : m_pDirEntries[0].dwLowStreamSize;
  m_nEffectiveNumMiniSectors = (size_t)(nRootEntrySize / m_nMiniSectorSize);

  // �������� ������ ��������� ������� ��������
  res2 = CreateDirEntryItemList();
  if (res2 != O2_OK)
  {
    if ((res2 == O2_ERROR_MEM) || (m_nNumDirEntryItems == 0))
      return res2;
    if (res == O2_OK)
      res = res2;
  }

  m_bOpened = true;

  return res;
}
/***************************************************************************/
/* Close - ��������                                                        */
/***************************************************************************/
void COLE2File::Close()
{
  m_bOpened = false;
  m_pInStream = NULL;
  m_nMajorVersion = 0;
  m_nMinorVersion = 0;
  m_nEffectiveNumSectors = 0;
  m_nEffectiveNumMiniSectors = 0;
  m_dwRootEntryStartSector = ENDOFCHAIN;
  m_nSectorSize = 0;
  m_nMiniSectorSize = 0;
  m_nMiniStreamCutoffSize = 0;
  m_nNumFATEntries = 0;
  m_nNumMiniFATEntries = 0;
  m_nNumDirEntries = 0;
  m_bSomeDirEntriesLost = false;
  m_nFileSize = 0i64;
  m_nNumDirEntryItems = 0;
  m_pDirEntryNameAliases = NULL;
  m_nNumDirEntryNameAliases = 0;
  memset(&m_fileHeader, 0, COMPOUND_FILE_HEADER_SIZE);
  // ����������� ������ ��� ���������� �������� "����������" �������
  if (m_pLostStorageDirEntry)
  {
    free(m_pLostStorageDirEntry);
    m_pLostStorageDirEntry = NULL;
  }
  // ����������� ������ ��������� ������� ��������
  if (m_pDirEntryItems)
  {
    free(m_pDirEntryItems);
    m_pDirEntryItems = NULL;
  }
  // ����������� ������ ������� ��������
  if (m_pDirEntries)
  {
    free(m_pDirEntries);
    m_pDirEntries = NULL;
  }
  // ����������� mini FAT
  if (m_pMiniFAT)
  {
    free(m_pMiniFAT);
    m_pMiniFAT = NULL;
  }
  // ����������� FAT
  if (m_pFAT)
  {
    free(m_pFAT);
    m_pFAT = NULL;
  }
}
/***************************************************************************/
/* GetDirEntryItem - ��������� �������� ������ ��������                    */
/***************************************************************************/
const DirEntryItem *COLE2File::GetDirEntryItem(
  size_t nIndex
  ) const
{
  if (nIndex < m_nNumDirEntryItems)
    return &m_pDirEntryItems[nIndex];
  return NULL;
}
/***************************************************************************/
/* GetDirEntry - ��������� ������ ��������                                 */
/***************************************************************************/
const DIRECTORY_ENTRY *COLE2File::GetDirEntry(
  uint32_t dwID
  ) const
{
  if (dwID < m_nNumDirEntries)
    return &m_pDirEntries[dwID];
  if (m_bSomeDirEntriesLost && (dwID == LOSTSTORAGEID))
    return m_pLostStorageDirEntry;
  return NULL;
}
/***************************************************************************/
/* SetDirEntryNameAliases - ��������� ������ ����������� ���� �������      */
/*                          ��������                                       */
/***************************************************************************/
void COLE2File::SetDirEntryNameAliases(
  const DirEntryNameAlias *pDirEntryNameAliases,
  size_t nNumDirEntryNameAliases
  )
{
  if (pDirEntryNameAliases && (nNumDirEntryNameAliases != 0))
  {
    m_pDirEntryNameAliases = pDirEntryNameAliases;
    m_nNumDirEntryNameAliases = nNumDirEntryNameAliases;
  }
  else
  {
    m_pDirEntryNameAliases = NULL;
    m_nNumDirEntryNameAliases = 0;
  }
}
/***************************************************************************/
/* FindDirEntryNameAlias - ����� ���������� ����� ������ ��������          */
/***************************************************************************/
const wchar_t *COLE2File::FindDirEntryNameAlias(
  uint32_t dwDirEntryID
  ) const
{
  if (!m_pDirEntryNameAliases || (m_nNumDirEntryNameAliases == 0))
    return NULL;
  for (size_t i = 0; i < m_nNumDirEntryNameAliases; i++)
  {
    if (dwDirEntryID == m_pDirEntryNameAliases[i].dwDirEntryID)
      return m_pDirEntryNameAliases[i].wszNameAlias;
  }
  return NULL;
}
/***************************************************************************/
/* GetDirEntryByPath - ��������� ������ �������� �� ����                   */
/***************************************************************************/
uint32_t COLE2File::GetDirEntryByPath(
  const wchar_t *pwszPath,
  uint32_t dwBaseStorageID,
  bool bUseAliases
  ) const
{
  if (!pwszPath || !pwszPath[0] ||
      (m_nNumDirEntryItems == 0))
    return NOSTREAM;

  uint32_t dwDirEntryID;

  const wchar_t *pwch = pwszPath;

  if (ISPATHDELIMITER(*pwch))
  {
    pwch++;
    if (*pwch == L'\0')
      return 0;
    dwDirEntryID = 0;
  }
  else
  {
    if (dwBaseStorageID == NOSTREAM)
      return NOSTREAM;
    dwDirEntryID = dwBaseStorageID;
  }

  do
  {
    const wchar_t *pwchName = pwch;
    while ((*pwch != L'\0') && !ISPATHDELIMITER(*pwch))
      pwch++;
    size_t cchName = pwch - pwchName;
    if (cchName == 0)
      return NOSTREAM;
    if ((pwchName[0] == L'.') &&
        ((cchName == 1) || ((cchName == 2) && (pwchName[1] == L'.'))))
    {
      if (cchName == 2)
      {
        if (dwDirEntryID != 0)
        {
          // ��������� �������������� ������������ ������ ��������
          dwDirEntryID = GetParentDirEntryID(dwDirEntryID);
          if (dwDirEntryID == NOSTREAM)
            return NOSTREAM;
        }
      }
    }
    else
    {
      // ����� ������ ��������
      dwDirEntryID = FindDirEntry(dwDirEntryID, 0, pwchName, cchName,
                                  bUseAliases);
      if (dwDirEntryID == NOSTREAM)
        return NOSTREAM;
      if (*pwch != L'\0')
      {
        // ����������?
        const DIRECTORY_ENTRY *pDirEntry = GetDirEntry(dwDirEntryID);
        if (!pDirEntry ||
            ((pDirEntry->bObjectType != OBJ_TYPE_STORAGE) &&
             (pDirEntry->bObjectType != OBJ_TYPE_ROOT)))
          return NOSTREAM;
      }
    }
    if (*pwch != L'\0')
      pwch++;
  }
  while (*pwch != L'\0');

  return dwDirEntryID;
}
/***************************************************************************/
/* FindDirEntry - ����� ������ ��������                                    */
/***************************************************************************/
uint32_t COLE2File::FindDirEntry(
  uint32_t dwStorageID,
  uint8_t bObjectType,
  const wchar_t *pwszName,
  bool bUseAliases
  ) const
{
  size_t cchName;
#ifdef USE_WINAPI
  cchName = ::lstrlen(pwszName);
#else
  cchName = wcslen(pwszName);
#endif  // USE_WINAPI
  // ����� ������ ��������
  return FindDirEntry(dwStorageID, bObjectType, pwszName, cchName,
                      bUseAliases);
}
/***************************************************************************/
/* FindDirEntry - ����� ������ ��������                                    */
/***************************************************************************/
uint32_t COLE2File::FindDirEntry(
  uint32_t dwStorageID,
  uint8_t bObjectType,
  const wchar_t *pwchName,
  size_t cchName,
  bool bUseAliases
  ) const
{
  for (size_t i = 0; i < m_nNumDirEntryItems; i++)
  {
    const DirEntryItem *pDirEntryItem = &m_pDirEntryItems[i];
    if ((dwStorageID == pDirEntryItem->dwParentID) &&
        ((bObjectType == 0) ||
         (bObjectType == pDirEntryItem->pDirEntry->bObjectType)))
    {
      const wchar_t *pwchDirEntryName;
      size_t cchDirEntryName;
      // ����� ���������� ����� ������ ��������
      if (bUseAliases &&
          (pwchDirEntryName = FindDirEntryNameAlias(pDirEntryItem->dwID)) &&
          pwchDirEntryName[0])
      {
#ifdef USE_WINAPI
        cchDirEntryName = ::lstrlen(pwchDirEntryName);
#else
        cchDirEntryName = wcslen(pwchDirEntryName);
#endif  // USE_WINAPI
      }
      else
      {
        cchDirEntryName = pDirEntryItem->pDirEntry->wSizeOfName >> 1;
        if ((cchDirEntryName <= 1) ||
            (cchDirEntryName > (MAX_DIR_ENTRY_NAME_SIZE >> 1)))
          continue;
        cchDirEntryName--;
        pwchDirEntryName = (wchar_t *)pDirEntryItem->pDirEntry->bName;
      }
      if (cchName == cchDirEntryName)
      {
#ifdef USE_WINAPI
        if (CSTR_EQUAL == ::CompareStringW(LOCALE_SYSTEM_DEFAULT,
                                           NORM_IGNORECASE,
                                           pwchDirEntryName, (int)cchName,
                                           pwchName, (int)cchName))
#else
        if (!_wcsnicoll(pwchDirEntryName, pwchName, cchName))
#endif  // USE_WINAPI
        {
          return pDirEntryItem->dwID;
        }
      }
    }
  }
  return NOSTREAM;
}
/***************************************************************************/
/* GetParentDirEntryID - ��������� �������������� ������������ ������      */
/*                       ��������                                          */
/***************************************************************************/
uint32_t COLE2File::GetParentDirEntryID(
  uint32_t dwDirEntryID
  ) const
{
  for (size_t i = 0; i < m_nNumDirEntryItems; i++)
  {
    if (dwDirEntryID == m_pDirEntryItems[i].dwID)
      return m_pDirEntryItems[i].dwParentID;
  }
  return NOSTREAM;
}
/***************************************************************************/
/* GetDirEntryFullPath - ��������� ������� ���� ������ ��������            */
/***************************************************************************/
size_t COLE2File::GetDirEntryFullPath(
  uint32_t dwDirEntryID,
  wchar_t *pBuffer,
  size_t nBufferSize,
  bool bUseAliases
  ) const
{
  // ��������� ������� ���� ������ ��������
  size_t nRet = RecurseGetDirEntryFullPath(dwDirEntryID, pBuffer,
                                           nBufferSize, bUseAliases);
  if (nRet == 1)
  {
    // �������� �������
    if ((nBufferSize >= 2) && pBuffer)
    {
      pBuffer[0] = PATH_DELIMITER;
      pBuffer[1] = L'\0';
    }
    else
    {
      nRet++;
    }
  }
  else
  {
    if ((nRet != 0) && (nRet <= nBufferSize))
      nRet--;
  }
  return nRet;
}
/***************************************************************************/
/* RecurseGetDirEntryFullPath - ����������� ������� ��������� ������� ���� */
/*                              ������ ��������                            */
/***************************************************************************/
size_t COLE2File::RecurseGetDirEntryFullPath(
  uint32_t dwDirEntryID,
  wchar_t *pBuffer,
  size_t nBufferSize,
  bool bUseAliases
  ) const
{
  if (dwDirEntryID == NOSTREAM)
    return 0;

  if (dwDirEntryID == 0)
  {
    if ((nBufferSize != 0) && pBuffer)
      pBuffer[0] = L'\0';
    return 1;
  }

  // ����������� ����� �������
  size_t nRet = RecurseGetDirEntryFullPath(GetParentDirEntryID(dwDirEntryID),
                                           pBuffer, nBufferSize,
                                           bUseAliases);
  if (nRet == 0)
    return 0;
  const DIRECTORY_ENTRY *pDirEntry = GetDirEntry(dwDirEntryID);
  if (pDirEntry == NULL)
    return 0;
  const wchar_t *pwchName;
  size_t cchName;
  // ����� ���������� ����� ������ ��������
  if (bUseAliases &&
      (pwchName = FindDirEntryNameAlias(dwDirEntryID)) &&
      pwchName[0])
  {
#ifdef USE_WINAPI
    cchName = ::lstrlen(pwchName);
#else
    cchName = wcslen(pwchName);
#endif  // USE_WINAPI
  }
  else
  {
    pwchName = (const wchar_t *)pDirEntry->bName;
    cchName = pDirEntry->wSizeOfName >> 1;
    if ((cchName <= 1) || (cchName > (MAX_DIR_ENTRY_NAME_SIZE >> 1)))
      return 0;
    cchName--;
  }
  if ((nRet + cchName + 1 <= nBufferSize) && pBuffer)
  {
    pBuffer[nRet - 1] = PATH_DELIMITER;
    memcpy(pBuffer + nRet, pwchName, cchName << 1);
    pBuffer[nRet + cchName] = L'\0';
  }
  return (nRet + cchName + 1);
}
/***************************************************************************/
/* ExtractStreamData - ���������� ����������� ������                       */
/***************************************************************************/
O2Res COLE2File::ExtractStreamData(
  uint32_t dwStreamID,
  uint64_t nOffset,
  size_t nSize,
  void *pBuffer,
  size_t *pnProcessedSize
  ) const
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  const DIRECTORY_ENTRY *pDirEntry = GetDirEntry(dwStreamID);
  if ((!pDirEntry) ||
      (pDirEntry->bObjectType != OBJ_TYPE_STREAM) ||
      !pBuffer)
    return O2_ERROR_PARAM;

  // ���������� ����������� ������
  return ExtractStreamData(pDirEntry->dwStartSector,
                           (m_nMajorVersion > 3)
                             ? pDirEntry->qwStreamSize
                             : pDirEntry->dwLowStreamSize,
                           nOffset, nSize, pBuffer, pnProcessedSize);
}
/***************************************************************************/
/* ExtractStreamData - ���������� ����������� ������                       */
/***************************************************************************/
O2Res COLE2File::ExtractStreamData(
  uint32_t dwStreamID,
  uint64_t nOffset,
  uint64_t nSize,
  CSeqOutStream *pOutStream,
  uint64_t *pnProcessedSize
  ) const
{
  if (pnProcessedSize)
    *pnProcessedSize = 0i64;

  const DIRECTORY_ENTRY *pDirEntry = GetDirEntry(dwStreamID);
  if ((!pDirEntry) ||
      (pDirEntry->bObjectType != OBJ_TYPE_STREAM) ||
      !pOutStream)
    return O2_ERROR_PARAM;

  // ���������� ����������� ������
  return ExtractStreamData(pDirEntry->dwStartSector,
                           (m_nMajorVersion > 3)
                             ? pDirEntry->qwStreamSize
                             : pDirEntry->dwLowStreamSize,
                           nOffset, nSize, pOutStream, pnProcessedSize);
}
/***************************************************************************/
/* ExtractStream - ���������� ������                                       */
/***************************************************************************/
O2Res COLE2File::ExtractStream(
  uint32_t dwStreamID,
  CSeqOutStream *pOutStream,
  uint64_t *pnProcessedSize
  ) const
{
  if (pnProcessedSize)
    *pnProcessedSize = 0i64;

  const DIRECTORY_ENTRY *pDirEntry = GetDirEntry(dwStreamID);
  if ((!pDirEntry) ||
      (pDirEntry->bObjectType != OBJ_TYPE_STREAM) ||
      !pOutStream)
    return O2_ERROR_PARAM;

  // ���������� ����������� ������
  uint64_t nStreamSize = (m_nMajorVersion > 3) ? pDirEntry->qwStreamSize
                                               : pDirEntry->dwLowStreamSize;
  return ExtractStreamData(pDirEntry->dwStartSector, nStreamSize, 0,
                           nStreamSize, pOutStream, pnProcessedSize);
}
/***************************************************************************/
/* CheckFileHeader - �������� ��������� �����                              */
/***************************************************************************/
bool COLE2File::CheckFileHeader() const
{
  return ((m_fileHeader.wSectorShift >= 0x0009) &&
          (m_fileHeader.wSectorShift <= 0x000C) &&
          (m_fileHeader.wMiniSectorShift > 0) &&
          (m_fileHeader.wMiniSectorShift < m_fileHeader.wSectorShift));
}
/***************************************************************************/
/* ExtractStreamData - ���������� ����������� ������                       */
/***************************************************************************/
O2Res COLE2File::ExtractStreamData(
  uint32_t dwStreamStartSector,
  uint64_t nStreamSize,
  uint64_t nOffset,
  size_t nSize,
  void *pBuffer,
  size_t *pnProcessedSize
  ) const
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  if ((nStreamSize == 0i64) ||
      (nSize == 0))
    return O2_OK;

  if (nOffset >= nStreamSize)
    return O2_ERROR_PARAM;

  if (nSize > nStreamSize - nOffset)
    nSize = (size_t)(nStreamSize - nOffset);

  O2Res res = O2_OK;
  uint8_t *pb = (uint8_t *)pBuffer;
  uint32_t dwSectorNum = dwStreamStartSector;
  size_t nOffsetInSector = 0;

  size_t nBytesToRead;
  size_t nBytesRead;

  if (nStreamSize < m_nMiniStreamCutoffSize)
  {
    if (nOffset != 0)
    {
      // ��������� ������ ���������� mini-������� � �������� ������ � ���
      if (!GetNextMiniSector(dwSectorNum,
                             (size_t)(nOffset / m_nMiniSectorSize),
                             &dwSectorNum))
        return O2_ERROR_DATA;
      nOffsetInSector = nOffset % m_nMiniSectorSize;
    }
    // ������ mini-�������� � �����
    while ((nSize != 0) && IsMiniSectorNumValid(dwSectorNum))
    {
      // ������ ������ mini-������� � �����
      nBytesToRead = min(m_nMiniSectorSize - nOffsetInSector, nSize);
      res = ReadMiniSectorData(dwSectorNum, nOffsetInSector, nBytesToRead,
                               pb, &nBytesRead);
      pb += nBytesRead;
      if (res != O2_OK)
        break;
      nOffsetInSector = 0;
      nSize -= nBytesRead;
      dwSectorNum = m_pMiniFAT[dwSectorNum];
    }
  }
  else
  {
    if (nOffset != 0)
    {
      // ��������� ������ ���������� ������� � �������� ������ � ���
      if (!GetNextSector(dwSectorNum, (size_t)(nOffset / m_nSectorSize),
                         &dwSectorNum))
        return O2_ERROR_DATA;
      nOffsetInSector = nOffset % m_nSectorSize;
    }
    // ������ �������� � �����
    while ((nSize != 0) && IsSectorNumValid(dwSectorNum))
    {
      // ������ ������ ������� � �����
      nBytesToRead = min(m_nSectorSize - nOffsetInSector, nSize);
      res = ReadSectorData(dwSectorNum, nOffsetInSector, nBytesToRead, pb,
                           &nBytesRead);
      pb += nBytesRead;
      if (res != O2_OK)
        break;
      nOffsetInSector = 0;
      nSize -= nBytesRead;
      dwSectorNum = m_pFAT[dwSectorNum];
    }
  }
  if (pnProcessedSize)
    *pnProcessedSize = pb - (uint8_t *)pBuffer;
  if ((nSize != 0) && (res == O2_OK))
    res = O2_ERROR_DATA;
  return res;
}
/***************************************************************************/
/* ExtractStreamData - ���������� ������ ������                            */
/***************************************************************************/
O2Res COLE2File::ExtractStreamData(
  uint32_t dwStreamStartSector,
  uint64_t nStreamSize,
  uint64_t nOffset,
  uint64_t nSize,
  CSeqOutStream *pOutStream,
  uint64_t *pnProcessedSize
  ) const
{
  if (pnProcessedSize)
    *pnProcessedSize = 0i64;

  if ((nStreamSize == 0i64) ||
      (nSize == 0i64))
    return O2_OK;

  if (nOffset >= nStreamSize)
    return O2_ERROR_PARAM;

  if (nSize > nStreamSize - nOffset)
    nSize = nStreamSize - nOffset;

  O2Res res = O2_OK;
  uint32_t dwSectorNum = dwStreamStartSector;
  size_t nOffsetInSector = 0;

  size_t nBytesToRead;
  size_t nBytesRead;
  size_t nBytesWritten;
  void *pSectorBuffer;

  if (nStreamSize < m_nMiniStreamCutoffSize)
  {
    if (nOffset != 0)
    {
      // ��������� ������ ���������� mini-������� � �������� ������ � ���
      if (!GetNextMiniSector(dwSectorNum,
                             (size_t)(nOffset / m_nMiniSectorSize),
                             &dwSectorNum))
        return O2_ERROR_DATA;
      nOffsetInSector = nOffset % m_nMiniSectorSize;
    }
    // ������ mini-�������� � ������ �� � ����� ��� ������ ������
    pSectorBuffer = _alloca(m_nMiniSectorSize);
    while ((nSize != 0i64) && IsMiniSectorNumValid(dwSectorNum))
    {
      // ������ ������ mini-������� � �����
      nBytesToRead = m_nMiniSectorSize - nOffsetInSector;
      if (nBytesToRead > nSize) nBytesToRead = (size_t)nSize;
      res = ReadMiniSectorData(dwSectorNum, nOffsetInSector, nBytesToRead,
                               pSectorBuffer, &nBytesRead);
      if (nBytesRead != 0)
      {
        // ������ ������ � �����
        nBytesWritten = pOutStream->Write(pSectorBuffer, nBytesRead);
        if (pnProcessedSize)
          *pnProcessedSize += nBytesWritten;
        if (nBytesWritten != nBytesRead)
          res = O2_ERROR_WRITE;
      }
      if (res != O2_OK)
        break;
      nOffsetInSector = 0;
      nSize -= nBytesRead;
      dwSectorNum = m_pMiniFAT[dwSectorNum];
    }
  }
  else
  {
    if (nOffset != 0)
    {
      // ��������� ������ ���������� ������� � �������� ������ � ���
      if (!GetNextSector(dwSectorNum, (size_t)(nOffset / m_nSectorSize),
                         &dwSectorNum))
        return O2_ERROR_DATA;
      nOffsetInSector = nOffset % m_nSectorSize;
    }
    // ������ �������� � ������ �� � ����� ��� ������ ������
    pSectorBuffer = _alloca(m_nSectorSize);
    while ((nSize != 0i64) && IsSectorNumValid(dwSectorNum))
    {
      // ������ ������ ������� � �����
      nBytesToRead = m_nSectorSize - nOffsetInSector;
      if (nBytesToRead > nSize) nBytesToRead = (size_t)nSize;
      res = ReadSectorData(dwSectorNum, nOffsetInSector, nBytesToRead,
                           pSectorBuffer, &nBytesRead);
      if (nBytesRead != 0)
      {
        nBytesWritten = pOutStream->Write(pSectorBuffer, nBytesRead);
        if (pnProcessedSize)
          *pnProcessedSize += nBytesWritten;
        if (nBytesWritten != nBytesRead)
          res = O2_ERROR_WRITE;
      }
      if (res != O2_OK)
        break;
      nOffsetInSector = 0;
      nSize -= nBytesRead;
      dwSectorNum = m_pFAT[dwSectorNum];
    }
  }
  if ((nSize != 0i64) && (res == O2_OK))
    res = O2_ERROR_DATA;
  return res;
}
/***************************************************************************/
/* AllocAndReadFAT - ��������� ������ � ������ FAT                         */
/***************************************************************************/
O2Res COLE2File::AllocAndReadFAT()
{
  size_t nSectorCount = m_fileHeader.dwNumFATSectors;
  if (nSectorCount == 0)
    return O2_ERROR_DATA;

  m_pFAT = (uint32_t *)malloc(nSectorCount * m_nSectorSize);
  if (!m_pFAT)
    return O2_ERROR_MEM;

  O2Res res = O2_OK;
  uint8_t *pb = (uint8_t *)m_pFAT;

  uint32_t *pdwSectorNum;
  size_t i;
  size_t nBytesRead;

  // ������ �������� FAT �� ������� DIFAT � ��������� �����
  pdwSectorNum = m_fileHeader.dwDIFAT;
  for (i = min(NUMBER_OF_DIFAT_ENTRIES, nSectorCount);
       i != 0;
       i--, nSectorCount--, pdwSectorNum++)
  {
    if (*pdwSectorNum >= MAXREGSECT)
    {
      res = O2_ERROR_DATA;
      break;
    }
    // ������ ������� FAT
    res = ReadSector(*pdwSectorNum, pb, &nBytesRead);
    pb += nBytesRead;
    if (res != O2_OK)
      break;
  }

  if ((nSectorCount != 0) && (res == O2_OK))
  {
    // ������ �������� FAT, ��������� � ����������� DIFAT
    uint32_t *pdwDIFATSectors = (uint32_t *)_alloca(m_nSectorSize);
    pdwSectorNum = &m_fileHeader.dwDIFATStartSector;
    i = m_fileHeader.dwNumDIFATSectors;
    for (; (i != 0) && (nSectorCount != 0); i--)
    {
      if (*pdwSectorNum >= MAXREGSECT)
      {
        res = O2_ERROR_DATA;
        break;
      }
      // ������ ������� ����������� DIFAT
      res = ReadSector(*pdwSectorNum, pdwDIFATSectors, &nBytesRead);
      size_t j = nBytesRead / sizeof(uint32_t);
      if (res == O2_OK) j--;
      if (j > nSectorCount) j = nSectorCount;
      pdwSectorNum = pdwDIFATSectors;
      for (; j != 0; j--, nSectorCount--, pdwSectorNum++)
      {
        if (*pdwSectorNum >= MAXREGSECT)
        {
          res = O2_ERROR_DATA;
          break;
        }
        // ������ ������� FAT
        O2Res res2 = ReadSector(*pdwSectorNum, pb, &nBytesRead);
        pb += nBytesRead;
        if (res2 != O2_OK)
        {
          res = res2;
          break;
        }
      }
      if (res != O2_OK)
        break;
    }
    if (((i != 0) || (nSectorCount != 0)) && (res == O2_OK))
      res = O2_ERROR_DATA;
  }

  // ����������� ���������� ������� � FAT
  m_nNumFATEntries = (size_t)((pb - (uint8_t *)m_pFAT) / sizeof(uint32_t));
  // ���������� �������������� �������
  while ((m_nNumFATEntries != 0) &&
         (m_pFAT[m_nNumFATEntries - 1] == FREESECT))
  {
    m_nNumFATEntries--;
  }

  if ((m_nNumFATEntries == 0) && (res == O2_OK))
    res = O2_ERROR_DATA;
  return res;
}
/***************************************************************************/
/* AllocAndReadMiniFAT - ��������� ������ � ������ mini FAT                */
/***************************************************************************/
O2Res COLE2File::AllocAndReadMiniFAT()
{
  O2Res res = O2_OK;

  size_t nSectorCount;
  uint32_t dwSectorNum;

  // ����������� ���������� �������� mini FAT � �����
  nSectorCount = 0;
  dwSectorNum = m_fileHeader.dwMiniFATStartSector;
  while ((dwSectorNum != ENDOFCHAIN) && (nSectorCount < m_nNumFATEntries))
  {
    if (!IsSectorNumValid(dwSectorNum))
      break;
    dwSectorNum = m_pFAT[dwSectorNum];
    nSectorCount++;
  }
  if (dwSectorNum != ENDOFCHAIN)
    res = O2_ERROR_DATA;
  if (nSectorCount == 0)
    return res;

  m_pMiniFAT = (uint32_t *)malloc(nSectorCount * m_nSectorSize);
  if (!m_pMiniFAT)
    return O2_ERROR_MEM;

  uint8_t *pb = (uint8_t *)m_pMiniFAT;
  dwSectorNum = m_fileHeader.dwMiniFATStartSector;
  do
  {
    // ������ ������� � �����
    size_t nBytesRead;
    O2Res res2 = ReadSector(dwSectorNum, pb, &nBytesRead);
    pb += nBytesRead;
    if (res2 != O2_OK)
    {
      res = res2;
      break;
    }
    dwSectorNum = m_pFAT[dwSectorNum];
    nSectorCount--;
  }
  while (nSectorCount != 0);

  // ����������� ���������� ������� � mini FAT
  m_nNumMiniFATEntries = (size_t)((pb - (uint8_t *)m_pMiniFAT) /
                                  sizeof(uint32_t));
  // ���������� �������������� �������
  while ((m_nNumMiniFATEntries != 0) &&
         (m_pMiniFAT[m_nNumMiniFATEntries - 1] == FREESECT))
  {
    m_nNumMiniFATEntries--;
  }

  return res;
}
/***************************************************************************/
/* AllocAndReadDirectory - ��������� ������ � ������ ��������              */
/***************************************************************************/
O2Res COLE2File::AllocAndReadDirectory()
{
  O2Res res = O2_OK;

  size_t nSectorCount;
  uint32_t dwSectorNum;

  // ����������� ���������� �������� �������� � �����
  nSectorCount = 0;
  dwSectorNum = m_fileHeader.dwDirStartSector;
  while ((dwSectorNum != ENDOFCHAIN) && (nSectorCount < m_nNumFATEntries))
  {
    if (!IsSectorNumValid(dwSectorNum))
      break;
    dwSectorNum = m_pFAT[dwSectorNum];
    nSectorCount++;
  }
  if (nSectorCount == 0)
    return O2_ERROR_DATA;
  if (dwSectorNum != ENDOFCHAIN)
    res = O2_ERROR_DATA;

  m_pDirEntries = (PDIRECTORY_ENTRY)malloc(nSectorCount * m_nSectorSize);
  if (!m_pDirEntries)
    return O2_ERROR_MEM;

  uint8_t *pb = (uint8_t *)m_pDirEntries;
  dwSectorNum = m_fileHeader.dwDirStartSector;
  do
  {
    // ������ ������� � �����
    size_t nBytesRead;
    O2Res res2 = ReadSector(dwSectorNum, pb, &nBytesRead);
    pb += nBytesRead;
    if (res2 != O2_OK)
    {
      res = res2;
      break;
    }
    dwSectorNum = m_pFAT[dwSectorNum];
    nSectorCount--;
  }
  while (nSectorCount != 0);

  // ����������� ���������� ������� � ��������
  m_nNumDirEntries = (size_t)((pb - (uint8_t *)m_pDirEntries) /
                              DIRECTORY_ENTRY_SIZE);
  while ((m_nNumDirEntries != 0) &&
         (m_pDirEntries[m_nNumDirEntries - 1].bObjectType ==
            OBJ_TYPE_UNKNOWN))
  {
    m_nNumDirEntries--;
  }

  if ((m_nNumDirEntries == 0) && (m_nNumDirEntries >= MAXREGSID))
    return (res != O2_OK) ? res : O2_ERROR_DATA;

  // ��������� ����-�������� � ����� ������������ ������� ��������
  for (size_t i = 0; i < m_nNumDirEntries; i++)
  {
    size_t cchDirEntryName = min(MAX_DIR_ENTRY_NAME_SIZE >> 1,
                                 m_pDirEntries[i].wSizeOfName >> 1);
    if (cchDirEntryName != 0) cchDirEntryName--;
    ((uint16_t *)&m_pDirEntries[i].bName)[cchDirEntryName] = (uint16_t)0;
  }

  return res;
}
/***************************************************************************/
/* CreateDirEntryItemList - �������� ������ ��������� ������� ��������     */
/***************************************************************************/
O2Res COLE2File::CreateDirEntryItemList()
{
  // ��������� ������ ��� ������ ��������� ������� ��������
  // � ��������������� �������� ��� ���������� �������� "����������" �������
  size_t nNumDirEntryItems = m_nNumDirEntries + 1;
  m_pDirEntryItems = (PDirEntryItem)malloc(nNumDirEntryItems *
                                           sizeof(DirEntryItem));
  if (!m_pDirEntryItems)
    return O2_ERROR_MEM;

  // ������������ ������� ��������
  EnumDirEntries(0, NOSTREAM);
  if (m_nNumDirEntryItems == 0)
    return O2_ERROR_DATA;

  m_bSomeDirEntriesLost = false;
  size_t nFirstLostDirEntryItemIndex;

  if (m_nNumDirEntryItems < m_nNumDirEntries)
  {
    nFirstLostDirEntryItemIndex = m_nNumDirEntryItems;
    // ����� � ���������� "����������" �������
    for (uint32_t dwDirEntryID = 0; dwDirEntryID < m_nNumDirEntries;
         dwDirEntryID++)
    {
      size_t nDirEntryItemIndex = 0;
      while (nDirEntryItemIndex < m_nNumDirEntryItems)
      {
        if (dwDirEntryID == m_pDirEntryItems[nDirEntryItemIndex].dwID)
          break;
        nDirEntryItemIndex++;
      }
      if (nDirEntryItemIndex >= m_nNumDirEntryItems)
      {
        m_bSomeDirEntriesLost = true;
        // ������������ ������� ��������
        EnumDirEntries(dwDirEntryID, LOSTSTORAGEID);
      }
    }
  }

  if (m_bSomeDirEntriesLost)
  {
    if (m_nNumDirEntryItems < nNumDirEntryItems)
    {
      // �������� ��������� ������ ��� �������� "����������" �������
      m_pLostStorageDirEntry =
        (PDIRECTORY_ENTRY)malloc(DIRECTORY_ENTRY_SIZE);
      if (!m_pLostStorageDirEntry)
        return O2_ERROR_MEM;
      memset(m_pLostStorageDirEntry, 0, DIRECTORY_ENTRY_SIZE);
      memcpy(m_pLostStorageDirEntry->bName, LOST_DIR_NAME,
             sizeof(LOST_DIR_NAME));
      m_pLostStorageDirEntry->wSizeOfName = sizeof(LOST_DIR_NAME);
      m_pLostStorageDirEntry->bObjectType = OBJ_TYPE_STORAGE;
      m_pLostStorageDirEntry->dwLeftSiblingID = NOSTREAM;
      m_pLostStorageDirEntry->dwRightSiblingID = NOSTREAM;
      m_pLostStorageDirEntry->dwChildID = NOSTREAM;

      PDirEntryItem pDirEntryItem = &m_pDirEntryItems[m_nNumDirEntryItems];
      pDirEntryItem->nType = DIR_ENTRY_TYPE_LOST;
      pDirEntryItem->dwID = LOSTSTORAGEID;
      pDirEntryItem->dwParentID = 0;
      pDirEntryItem->dwFirstChildID = NOSTREAM;
      // ����������� ������ �������� ������ �������� "����������" �������
      for (size_t i = nFirstLostDirEntryItemIndex; i < m_nNumDirEntryItems;
           i++)
      {
        if (m_pDirEntryItems[i].dwParentID == LOSTSTORAGEID)
        {
          pDirEntryItem->dwFirstChildID = m_pDirEntryItems[i].dwID;
          break;
        }
      }
      pDirEntryItem->pDirEntry = m_pLostStorageDirEntry;
      m_nNumDirEntryItems++;
    }
    return O2_ERROR_DATA;
  }

  return O2_OK;
}
/***************************************************************************/
/* EnumDirEntries - ����������� ������������ ������� ��������              */
/***************************************************************************/
void COLE2File::EnumDirEntries(
  uint32_t dwDirEntryID,
  uint32_t dwParentDirEntryID
  )
{
  const DIRECTORY_ENTRY *pDirEntry = GetDirEntry(dwDirEntryID);
  if (!pDirEntry || (pDirEntry->bObjectType == OBJ_TYPE_UNKNOWN))
    return;
  for (size_t i = 0; i < m_nNumDirEntryItems; i++)
  {
    if (dwDirEntryID == m_pDirEntryItems[i].dwID)
    {
      if (m_pDirEntryItems[i].dwParentID == LOSTSTORAGEID)
        m_pDirEntryItems[i].dwParentID = dwParentDirEntryID;
      return;
    }
  }

  PDirEntryItem pDirEntryItem = &m_pDirEntryItems[m_nNumDirEntryItems];
  pDirEntryItem->dwID = dwDirEntryID;
  pDirEntryItem->dwParentID = dwParentDirEntryID;
  if ((pDirEntry->bObjectType == OBJ_TYPE_STORAGE) ||
      (pDirEntry->bObjectType == OBJ_TYPE_ROOT))
  {
    pDirEntryItem->nType = DIR_ENTRY_TYPE_STORAGE;
    pDirEntryItem->dwFirstChildID = (pDirEntry->dwChildID != dwDirEntryID)
                                      ? pDirEntry->dwChildID
                                      : NOSTREAM;
  }
  else
  {
    pDirEntryItem->nType = (pDirEntry->bObjectType == OBJ_TYPE_STREAM)
                             ? DIR_ENTRY_TYPE_STREAM
                             : DIR_ENTRY_TYPE_UNKNOWN;
    pDirEntryItem->dwFirstChildID = NOSTREAM;
  }
  pDirEntryItem->pDirEntry = pDirEntry;
  if (++m_nNumDirEntryItems >= m_nNumDirEntries)
    return;

  // ����������� ������������ ��������� ��������
  // ������������� ������ �����
  if ((pDirEntry->dwLeftSiblingID != NOSTREAM) &&
      (pDirEntry->dwLeftSiblingID != dwDirEntryID))
  {
    EnumDirEntries(pDirEntry->dwLeftSiblingID, dwParentDirEntryID);
  }
  // ������������� ������ ������
  if ((pDirEntry->dwRightSiblingID != NOSTREAM) &&
      (pDirEntry->dwRightSiblingID != dwDirEntryID))
  {
    EnumDirEntries(pDirEntry->dwRightSiblingID, dwParentDirEntryID);
  }
  // �������� ������
  if (pDirEntryItem->dwFirstChildID != NOSTREAM)
  {
    EnumDirEntries(pDirEntryItem->dwFirstChildID, dwDirEntryID);
  }
}
/***************************************************************************/
/* ReadMiniSectorData - ������ ������ mini-������� � �����                 */
/***************************************************************************/
O2Res COLE2File::ReadMiniSectorData(
  uint32_t dwMiniSectorNum,
  size_t nOffset,
  size_t nBytesToRead,
  void *pBuffer,
  size_t *pnBytesRead
  ) const
{
  if (pnBytesRead)
    *pnBytesRead = 0;

  if ((nOffset >= m_nMiniSectorSize) ||
      (nBytesToRead > m_nMiniSectorSize - nOffset))
    return O2_ERROR_PARAM;

  // �������� ������������ ������ ������� mini FAT
  if (dwMiniSectorNum >= m_nEffectiveNumMiniSectors)
    return O2_ERROR_DATA;

  uint32_t dwSectorNum;

  // ����������� ������ ���������� �������
  if (!GetNextSector(m_dwRootEntryStartSector,
                     (dwMiniSectorNum * m_nMiniSectorSize) / m_nSectorSize,
                     &dwSectorNum) ||
      !IsSectorNumValid(dwSectorNum))
    return O2_ERROR_DATA;

  // ��������� �������� ������ � �������
  nOffset += (dwMiniSectorNum * m_nMiniSectorSize) % m_nSectorSize;

  // ������ ������ ������� � �����
  return ReadSectorData(dwSectorNum, nOffset, nBytesToRead, pBuffer,
                        pnBytesRead);
}
/***************************************************************************/
/* ReadSector - ������ ������� � �����                                     */
/***************************************************************************/
O2Res COLE2File::ReadSector(
  uint32_t dwSectorNum,
  void *pBuffer,
  size_t *pnBytesRead
  ) const
{
  // ������ ������ ������� � �����
  return ReadSectorData(dwSectorNum, 0, m_nSectorSize, pBuffer, pnBytesRead);
}
/***************************************************************************/
/* ReadSectorData - ������ ������ ������� � �����                          */
/***************************************************************************/
O2Res COLE2File::ReadSectorData(
  uint32_t dwSectorNum,
  size_t nOffset,
  size_t nBytesToRead,
  void *pBuffer,
  size_t *pnBytesRead
  ) const
{
  if (pnBytesRead)
    *pnBytesRead = 0;

  if ((nOffset >= m_nSectorSize) ||
      (nBytesToRead > m_nSectorSize - nOffset))
    return O2_ERROR_PARAM;

  uint64_t nSectorOffset;

  // ��������� �������� �������
  nSectorOffset = GetSectorOffset(dwSectorNum);
  if (nSectorOffset == (uint64_t)-1i64)
    return O2_ERROR_DATA;

  if (nBytesToRead == 0)
    return O2_OK;

  // ������ ������ ������
  return ReadStream(m_pInStream, nSectorOffset + nOffset, nBytesToRead,
                    pBuffer, pnBytesRead);
}
/***************************************************************************/
/* GetNextMiniSector - ��������� ������ ���������� mini-�������            */
/***************************************************************************/
bool COLE2File::GetNextMiniSector(
  uint32_t dwMiniSectorNum,
  size_t nIndex,
  uint32_t *pdwMiniSectorNum
  ) const
{
  while (nIndex != 0)
  {
    if (!IsMiniSectorNumValid(dwMiniSectorNum))
      return false;
    dwMiniSectorNum = m_pMiniFAT[dwMiniSectorNum];
    nIndex--;
  }
  *pdwMiniSectorNum = dwMiniSectorNum;
  return true;
}
/***************************************************************************/
/* GetNextSector - ��������� ������ ���������� �������                     */
/***************************************************************************/
bool COLE2File::GetNextSector(
  uint32_t dwSectorNum,
  size_t nIndex,
  uint32_t *pdwSectorNum
  ) const
{
  while (nIndex != 0)
  {
    if (!IsSectorNumValid(dwSectorNum))
      return false;
    dwSectorNum = m_pFAT[dwSectorNum];
    nIndex--;
  }
  *pdwSectorNum = dwSectorNum;
  return true;
}
/***************************************************************************/
/* IsMiniSectorNumValid - �������� ������������ ������ ������� mini FAT    */
/***************************************************************************/
bool COLE2File::IsMiniSectorNumValid(
  uint32_t dwMiniSectorNum
  ) const
{
  // m_nNumMiniFATEntries <= MAXREGSECT
  return (dwMiniSectorNum < m_nNumMiniFATEntries);
}
/***************************************************************************/
/* IsSectorNumValid - �������� ������������ ������ ������� FAT             */
/***************************************************************************/
bool COLE2File::IsSectorNumValid(
  uint32_t dwSectorNum
  ) const
{
  // m_nNumFATEntries <= MAXREGSECT
  return (dwSectorNum < m_nNumFATEntries);
}
/***************************************************************************/
/* GetSectorOffset - ��������� �������� �������                            */
/***************************************************************************/
uint64_t COLE2File::GetSectorOffset(
  uint32_t dwSectorNum
  ) const
{
  if (dwSectorNum < m_nEffectiveNumSectors)
    return ((dwSectorNum + 1) * m_nSectorSize);
  return (uint64_t)-1i64;
}
//---------------------------------------------------------------------------
