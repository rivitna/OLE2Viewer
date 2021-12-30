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
// Разделитель пути
#define PATH_DELIMITER      L'\\'
#define ISPATHDELIMITER(c)  ((c == L'\\') || (c == L'/'))

// Наименование фиктивного каталога для "потерянных" записей
// sizeof(LOST_DIR_NAME) <= MAX_DIR_ENTRY_NAME_SIZE
const wchar_t LOST_DIR_NAME[] = L"!LOST.DIR";
//---------------------------------------------------------------------------
/***************************************************************************/
/* ReadStream - Чтение данных потока                                       */
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

  // Установка указателя потока
  if (pInStream->Seek(&nOffset, STM_SEEK_SET))
    return O2_ERROR_DATA;

  if (nBytesToRead == 0)
    return O2_OK;

  O2Res res = O2_OK;

  // Чтение данных потока
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
/*                      COLE2File - класс файла OLE2                       */
/*                                                                         */
/***************************************************************************/

/***************************************************************************/
/* COLE2File - Конструктор класса                                          */
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
/* ~COLE2File - Деструктор класса                                          */
/***************************************************************************/
COLE2File::~COLE2File()
{
  // Закрытие
  Close();
}
/***************************************************************************/
/* Open - Открытие                                                         */
/***************************************************************************/
O2Res COLE2File::Open(
  CInStream *pInStream
  )
{
  // Закрытие
  Close();

  // Инициализация
  O2Res res = Init(pInStream);

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
O2Res COLE2File::Init(
  CInStream *pInStream
  )
{
  if (!pInStream)
    return O2_ERROR_PARAM;

  uint64_t nFileSize = 0i64;

  // Определение размера файла
  if (pInStream->Seek(&nFileSize, STM_SEEK_END) ||
      (nFileSize < COMPOUND_FILE_HEADER_SIZE))
    return O2_ERROR_DATA;

  O2Res res;

  // Чтение заголовка составного файла OLE2
  res = ReadStream(pInStream, 0i64, COMPOUND_FILE_HEADER_SIZE, &m_fileHeader,
                   NULL);
  if (res != O2_OK)
    return res;

  // Проверка заголовка составного файла OLE2
  if (!CheckFileHeader())
    return O2_ERROR_DATA;

  m_pInStream = pInStream;
  m_nFileSize = nFileSize;

  // Получение версии
  m_nMajorVersion = m_fileHeader.wMajorVersion;
  m_nMinorVersion = m_fileHeader.wMinorVersion;

  // Получение размеров секторов
  m_nSectorSize = (size_t)1 << m_fileHeader.wSectorShift;
  m_nMiniSectorSize = (size_t)1 << m_fileHeader.wMiniSectorShift;
  // Значение поля dwMiniStreamCutoffSize игнорируется
  m_nMiniStreamCutoffSize = 0x1000;  // m_fileHeader.dwMiniStreamCutoffSize;

  // Получение фактического количества секторов в файле
  m_nEffectiveNumSectors = (size_t)(m_nFileSize / m_nSectorSize);
  if (m_nEffectiveNumSectors <= 1)
    return O2_ERROR_DATA;
  m_nEffectiveNumSectors--;

  // Выделение памяти и чтение FAT
  res = AllocAndReadFAT();
  if ((res != O2_OK) && (m_nNumFATEntries == 0))
    return res;

  O2Res res2;

  // Выделение памяти и чтение mini FAT
  res2 = AllocAndReadMiniFAT();
  if ((res2 != O2_OK) && (res == O2_OK))
    res = res2;

  // Выделение памяти и чтение каталога
  res2 = AllocAndReadDirectory();
  if (res2 != O2_OK)
  {
    if ((m_nNumDirEntries == 0) && (m_nNumDirEntries >= MAXREGSID))
      return res2;
    if (res == O2_OK)
      res = res2;
  }

  // Получение номера начального сектора корневой записи
  m_dwRootEntryStartSector = m_pDirEntries[0].dwStartSector;
  // Получение фактического количества секторов в mini FAT
  uint64_t nRootEntrySize = (m_nMajorVersion > 3)
                              ? m_pDirEntries[0].qwStreamSize
                              : m_pDirEntries[0].dwLowStreamSize;
  m_nEffectiveNumMiniSectors = (size_t)(nRootEntrySize / m_nMiniSectorSize);

  // Создание списка элементов записей каталога
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
/* Close - Закрытие                                                        */
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
  // Уничтожение записи для фиктивного каталога "потерянных" записей
  if (m_pLostStorageDirEntry)
  {
    free(m_pLostStorageDirEntry);
    m_pLostStorageDirEntry = NULL;
  }
  // Уничтожение списка элементов записей каталога
  if (m_pDirEntryItems)
  {
    free(m_pDirEntryItems);
    m_pDirEntryItems = NULL;
  }
  // Уничтожение списка записей каталога
  if (m_pDirEntries)
  {
    free(m_pDirEntries);
    m_pDirEntries = NULL;
  }
  // Уничтожение mini FAT
  if (m_pMiniFAT)
  {
    free(m_pMiniFAT);
    m_pMiniFAT = NULL;
  }
  // Уничтожение FAT
  if (m_pFAT)
  {
    free(m_pFAT);
    m_pFAT = NULL;
  }
}
/***************************************************************************/
/* GetDirEntryItem - Получение элемента записи каталога                    */
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
/* GetDirEntry - Получение записи каталога                                 */
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
/* SetDirEntryNameAliases - Установка списка псевдонимов имен записей      */
/*                          каталога                                       */
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
/* FindDirEntryNameAlias - Поиск псевдонима имени записи каталога          */
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
/* GetDirEntryByPath - Получение записи каталога по пути                   */
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
          // Получение идентификатора родительской записи каталога
          dwDirEntryID = GetParentDirEntryID(dwDirEntryID);
          if (dwDirEntryID == NOSTREAM)
            return NOSTREAM;
        }
      }
    }
    else
    {
      // Поиск записи каталога
      dwDirEntryID = FindDirEntry(dwDirEntryID, 0, pwchName, cchName,
                                  bUseAliases);
      if (dwDirEntryID == NOSTREAM)
        return NOSTREAM;
      if (*pwch != L'\0')
      {
        // Подкаталог?
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
/* FindDirEntry - Поиск записи каталога                                    */
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
  // Поиск записи каталога
  return FindDirEntry(dwStorageID, bObjectType, pwszName, cchName,
                      bUseAliases);
}
/***************************************************************************/
/* FindDirEntry - Поиск записи каталога                                    */
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
      // Поиск псевдонима имени записи каталога
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
/* GetParentDirEntryID - Получение идентификатора родительской записи      */
/*                       каталога                                          */
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
/* GetDirEntryFullPath - Получение полного пути записи каталога            */
/***************************************************************************/
size_t COLE2File::GetDirEntryFullPath(
  uint32_t dwDirEntryID,
  wchar_t *pBuffer,
  size_t nBufferSize,
  bool bUseAliases
  ) const
{
  // Получение полного пути записи каталога
  size_t nRet = RecurseGetDirEntryFullPath(dwDirEntryID, pBuffer,
                                           nBufferSize, bUseAliases);
  if (nRet == 1)
  {
    // Корневой каталог
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
/* RecurseGetDirEntryFullPath - Рекурсивная функция получения полного пути */
/*                              записи каталога                            */
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

  // Рекурсивный вызов функции
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
  // Поиск псевдонима имени записи каталога
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
/* ExtractStreamData - Извлечение содержимого потока                       */
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

  // Извлечение содержимого потока
  return ExtractStreamData(pDirEntry->dwStartSector,
                           (m_nMajorVersion > 3)
                             ? pDirEntry->qwStreamSize
                             : pDirEntry->dwLowStreamSize,
                           nOffset, nSize, pBuffer, pnProcessedSize);
}
/***************************************************************************/
/* ExtractStreamData - Извлечение содержимого потока                       */
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

  // Извлечение содержимого потока
  return ExtractStreamData(pDirEntry->dwStartSector,
                           (m_nMajorVersion > 3)
                             ? pDirEntry->qwStreamSize
                             : pDirEntry->dwLowStreamSize,
                           nOffset, nSize, pOutStream, pnProcessedSize);
}
/***************************************************************************/
/* ExtractStream - Извлечение потока                                       */
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

  // Извлечение содержимого потока
  uint64_t nStreamSize = (m_nMajorVersion > 3) ? pDirEntry->qwStreamSize
                                               : pDirEntry->dwLowStreamSize;
  return ExtractStreamData(pDirEntry->dwStartSector, nStreamSize, 0,
                           nStreamSize, pOutStream, pnProcessedSize);
}
/***************************************************************************/
/* CheckFileHeader - Проверка заголовка файла                              */
/***************************************************************************/
bool COLE2File::CheckFileHeader() const
{
  return ((m_fileHeader.wSectorShift >= 0x0009) &&
          (m_fileHeader.wSectorShift <= 0x000C) &&
          (m_fileHeader.wMiniSectorShift > 0) &&
          (m_fileHeader.wMiniSectorShift < m_fileHeader.wSectorShift));
}
/***************************************************************************/
/* ExtractStreamData - Извлечение содержимого потока                       */
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
      // Получение номера начального mini-сектора и смещения данных в нем
      if (!GetNextMiniSector(dwSectorNum,
                             (size_t)(nOffset / m_nMiniSectorSize),
                             &dwSectorNum))
        return O2_ERROR_DATA;
      nOffsetInSector = nOffset % m_nMiniSectorSize;
    }
    // Чтение mini-секторов в буфер
    while ((nSize != 0) && IsMiniSectorNumValid(dwSectorNum))
    {
      // Чтение данных mini-сектора в буфер
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
      // Получение номера начального сектора и смещения данных в нем
      if (!GetNextSector(dwSectorNum, (size_t)(nOffset / m_nSectorSize),
                         &dwSectorNum))
        return O2_ERROR_DATA;
      nOffsetInSector = nOffset % m_nSectorSize;
    }
    // Чтение секторов в буфер
    while ((nSize != 0) && IsSectorNumValid(dwSectorNum))
    {
      // Чтение данных сектора в буфер
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
/* ExtractStreamData - Извлечение данных потока                            */
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
      // Получение номера начального mini-сектора и смещения данных в нем
      if (!GetNextMiniSector(dwSectorNum,
                             (size_t)(nOffset / m_nMiniSectorSize),
                             &dwSectorNum))
        return O2_ERROR_DATA;
      nOffsetInSector = nOffset % m_nMiniSectorSize;
    }
    // Чтение mini-секторов и запись их в поток для вывода данных
    pSectorBuffer = _alloca(m_nMiniSectorSize);
    while ((nSize != 0i64) && IsMiniSectorNumValid(dwSectorNum))
    {
      // Чтение данных mini-сектора в буфер
      nBytesToRead = m_nMiniSectorSize - nOffsetInSector;
      if (nBytesToRead > nSize) nBytesToRead = (size_t)nSize;
      res = ReadMiniSectorData(dwSectorNum, nOffsetInSector, nBytesToRead,
                               pSectorBuffer, &nBytesRead);
      if (nBytesRead != 0)
      {
        // Запись данных в поток
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
      // Получение номера начального сектора и смещения данных в нем
      if (!GetNextSector(dwSectorNum, (size_t)(nOffset / m_nSectorSize),
                         &dwSectorNum))
        return O2_ERROR_DATA;
      nOffsetInSector = nOffset % m_nSectorSize;
    }
    // Чтение секторов и запись их в поток для вывода данных
    pSectorBuffer = _alloca(m_nSectorSize);
    while ((nSize != 0i64) && IsSectorNumValid(dwSectorNum))
    {
      // Чтение данных сектора в буфер
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
/* AllocAndReadFAT - Выделение памяти и чтение FAT                         */
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

  // Чтение секторов FAT по записям DIFAT в заголовке файла
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
    // Чтение сектора FAT
    res = ReadSector(*pdwSectorNum, pb, &nBytesRead);
    pb += nBytesRead;
    if (res != O2_OK)
      break;
  }

  if ((nSectorCount != 0) && (res == O2_OK))
  {
    // Чтение секторов FAT, указанных в продолжении DIFAT
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
      // Чтение сектора продолжения DIFAT
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
        // Чтение сектора FAT
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

  // Определение количества записей в FAT
  m_nNumFATEntries = (size_t)((pb - (uint8_t *)m_pFAT) / sizeof(uint32_t));
  // Исключение неиспользуемых записей
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
/* AllocAndReadMiniFAT - Выделение памяти и чтение mini FAT                */
/***************************************************************************/
O2Res COLE2File::AllocAndReadMiniFAT()
{
  O2Res res = O2_OK;

  size_t nSectorCount;
  uint32_t dwSectorNum;

  // Определение количества секторов mini FAT в файле
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
    // Чтение сектора в буфер
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

  // Определение количества записей в mini FAT
  m_nNumMiniFATEntries = (size_t)((pb - (uint8_t *)m_pMiniFAT) /
                                  sizeof(uint32_t));
  // Исключение неиспользуемых записей
  while ((m_nNumMiniFATEntries != 0) &&
         (m_pMiniFAT[m_nNumMiniFATEntries - 1] == FREESECT))
  {
    m_nNumMiniFATEntries--;
  }

  return res;
}
/***************************************************************************/
/* AllocAndReadDirectory - Выделение памяти и чтение каталога              */
/***************************************************************************/
O2Res COLE2File::AllocAndReadDirectory()
{
  O2Res res = O2_OK;

  size_t nSectorCount;
  uint32_t dwSectorNum;

  // Определение количества секторов каталога в файле
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
    // Чтение сектора в буфер
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

  // Определение количества записей в каталоге
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

  // Установка нуль-символов в конце наименований записей каталога
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
/* CreateDirEntryItemList - Создание списка элементов записей каталога     */
/***************************************************************************/
O2Res COLE2File::CreateDirEntryItemList()
{
  // Выделение памяти для списка элементов записей каталога
  // с резервированием элемента для фиктивного каталога "потерянных" записей
  size_t nNumDirEntryItems = m_nNumDirEntries + 1;
  m_pDirEntryItems = (PDirEntryItem)malloc(nNumDirEntryItems *
                                           sizeof(DirEntryItem));
  if (!m_pDirEntryItems)
    return O2_ERROR_MEM;

  // Перечисление записей каталога
  EnumDirEntries(0, NOSTREAM);
  if (m_nNumDirEntryItems == 0)
    return O2_ERROR_DATA;

  m_bSomeDirEntriesLost = false;
  size_t nFirstLostDirEntryItemIndex;

  if (m_nNumDirEntryItems < m_nNumDirEntries)
  {
    nFirstLostDirEntryItemIndex = m_nNumDirEntryItems;
    // Поиск и добавление "потерянных" записей
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
        // Перечисление записей каталога
        EnumDirEntries(dwDirEntryID, LOSTSTORAGEID);
      }
    }
  }

  if (m_bSomeDirEntriesLost)
  {
    if (m_nNumDirEntryItems < nNumDirEntryItems)
    {
      // Создание фиктивной записи для каталога "потерянных" записей
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
      // Определение первой дочерней записи каталога "потерянных" записей
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
/* EnumDirEntries - Рекурсивное перечисление записей каталога              */
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

  // Рекурсивное перечисление связанных объектов
  // Одноуровневый объект слева
  if ((pDirEntry->dwLeftSiblingID != NOSTREAM) &&
      (pDirEntry->dwLeftSiblingID != dwDirEntryID))
  {
    EnumDirEntries(pDirEntry->dwLeftSiblingID, dwParentDirEntryID);
  }
  // Одноуровневый объект справа
  if ((pDirEntry->dwRightSiblingID != NOSTREAM) &&
      (pDirEntry->dwRightSiblingID != dwDirEntryID))
  {
    EnumDirEntries(pDirEntry->dwRightSiblingID, dwParentDirEntryID);
  }
  // Дочерний объект
  if (pDirEntryItem->dwFirstChildID != NOSTREAM)
  {
    EnumDirEntries(pDirEntryItem->dwFirstChildID, dwDirEntryID);
  }
}
/***************************************************************************/
/* ReadMiniSectorData - Чтение данных mini-сектора в буфер                 */
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

  // Проверка правильности номера сектора mini FAT
  if (dwMiniSectorNum >= m_nEffectiveNumMiniSectors)
    return O2_ERROR_DATA;

  uint32_t dwSectorNum;

  // Определение номера начального сектора
  if (!GetNextSector(m_dwRootEntryStartSector,
                     (dwMiniSectorNum * m_nMiniSectorSize) / m_nSectorSize,
                     &dwSectorNum) ||
      !IsSectorNumValid(dwSectorNum))
    return O2_ERROR_DATA;

  // Получение смещения данных в секторе
  nOffset += (dwMiniSectorNum * m_nMiniSectorSize) % m_nSectorSize;

  // Чтение данных сектора в буфер
  return ReadSectorData(dwSectorNum, nOffset, nBytesToRead, pBuffer,
                        pnBytesRead);
}
/***************************************************************************/
/* ReadSector - Чтение сектора в буфер                                     */
/***************************************************************************/
O2Res COLE2File::ReadSector(
  uint32_t dwSectorNum,
  void *pBuffer,
  size_t *pnBytesRead
  ) const
{
  // Чтение данных сектора в буфер
  return ReadSectorData(dwSectorNum, 0, m_nSectorSize, pBuffer, pnBytesRead);
}
/***************************************************************************/
/* ReadSectorData - Чтение данных сектора в буфер                          */
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

  // Получение смещение сектора
  nSectorOffset = GetSectorOffset(dwSectorNum);
  if (nSectorOffset == (uint64_t)-1i64)
    return O2_ERROR_DATA;

  if (nBytesToRead == 0)
    return O2_OK;

  // Чтение данных потока
  return ReadStream(m_pInStream, nSectorOffset + nOffset, nBytesToRead,
                    pBuffer, pnBytesRead);
}
/***************************************************************************/
/* GetNextMiniSector - Получение номера следующего mini-сектора            */
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
/* GetNextSector - Получение номера следующего сектора                     */
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
/* IsMiniSectorNumValid - Проверка правильности номера сектора mini FAT    */
/***************************************************************************/
bool COLE2File::IsMiniSectorNumValid(
  uint32_t dwMiniSectorNum
  ) const
{
  // m_nNumMiniFATEntries <= MAXREGSECT
  return (dwMiniSectorNum < m_nNumMiniFATEntries);
}
/***************************************************************************/
/* IsSectorNumValid - Проверка правильности номера сектора FAT             */
/***************************************************************************/
bool COLE2File::IsSectorNumValid(
  uint32_t dwSectorNum
  ) const
{
  // m_nNumFATEntries <= MAXREGSECT
  return (dwSectorNum < m_nNumFATEntries);
}
/***************************************************************************/
/* GetSectorOffset - Получение смещение сектора                            */
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
