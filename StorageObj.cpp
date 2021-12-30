//---------------------------------------------------------------------------
#define WIN32_LEAN_AND_MEAN  // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#include <malloc.h>
#include <memory.h>
#include "StrUtils.h"
#include "FileUtil.h"
#include "OLE2Type.h"
#include "OLE2CF.h"
#include "OLE2Util.h"
#include "OLE2MSI.h"
#include "OLE2Metadata.h"
#include "OLE2VBA.h"
#include "OLE2Word.h"
#include "OLE2Native.h"
#include "StorageObj.h"
//---------------------------------------------------------------------------
#ifndef countof
#define countof(a)  (sizeof(a)/sizeof(a[0]))
#endif  // countof
//---------------------------------------------------------------------------
// ������ ������ �� ��������� ��� ���� � �������� ��������
#define DEFAULT_CURRENTDIRPATH_LEN  256
//---------------------------------------------------------------------------
// ��� ����� � ������� �� ��������� ������ ��������� Word
const wchar_t WORDDOC_TEXT_FILE_NAME[] = L"{text}";
// ��� ���������� ����� � �����������
const wchar_t INFO_TEXT_FILE_NAME[]    = L"{info}.txt";

// ������������ ����������� �������
const wchar_t* const OPENABLE_STREAM_NAMES[] =
{
  L"WordDocument",
  L"\5SummaryInformation",
  L"\5DocumentSummaryInformation",
  L"\1Ole10Native"
};

// ������������ ����� ������ OLE2
const wchar_t* const OLE2_TYPE_NAMES[] =
{
  L"OLE2 Unknown",
  L"Microsoft Word",
  L"Microsoft Excel",
  L"Microsoft PowerPoint",
  L"Microsoft Visio",
  L"Microsoft Excel 5",
  L"Ichitaro",
  L"MSI/MSP/MSM"
};

// ��� ������������� ��������
const wchar_t PARENT_ITEM_NAME[] = L"..";
//---------------------------------------------------------------------------
/***************************************************************************/
/* OLE2FileTypeToString - �������������� ���� ����� OLE2 � ������          */
/***************************************************************************/
const wchar_t *OLE2FileTypeToString(
  int nFileType
  )
{
  if (nFileType >= countof(OLE2_TYPE_NAMES))
    nFileType = 0;
  return OLE2_TYPE_NAMES[nFileType];
}
//---------------------------------------------------------------------------
/***************************************************************************/
/*                                                                         */
/*                 CStorageObject - ����� �����-����������                 */
/*                                                                         */
/***************************************************************************/

/***************************************************************************/
/* CStorageObject - ����������� ������                                     */
/***************************************************************************/
CStorageObject::CStorageObject()
  : UseOpenableStreamAttr(false),
    m_pwszFilePath(NULL),
    m_pwszFileName(NULL),
    m_nFileType(OLE2_UNKNOWN),
    m_nNumStreams(0),
    m_nNumStorages(0),
    m_nNumMacros(0),
    m_nFindItemIndex(0),
    m_dwCurParentItemID(0),
    m_bCurParentItemIsStream(false),
    m_cchCurrentDirPathBuf(0),
    m_pwszCurrentDirPath(NULL),
    m_pItemNameAliases(NULL),
    m_nNumItemNameAliases(0),
    m_inStream(),
    m_ole2File()
{
}
/***************************************************************************/
/* CStorageObject - ���������� ������                                      */
/***************************************************************************/
CStorageObject::~CStorageObject()
{
  // ��������
  Close();
}
/***************************************************************************/
/* Open - ��������                                                        */
/***************************************************************************/
int CStorageObject::Open(
  const wchar_t *pwszFileName
  )
{
  // ��������
  Close();

  // �������������
  int res = Init(pwszFileName);

  if (!GetOpened())
  {
    // ��������
    Close();
  }

  return res;
}
/***************************************************************************/
/* Init - �������������                                                    */
/***************************************************************************/
int CStorageObject::Init(
  const wchar_t *pwszFileName
  )
{
  // �������� �����
  if (m_inStream.Open(pwszFileName, true))
    return O2_ERROR_OPEN;

  // �������� ��������� ����� OLE2
  uint32_t dwSign[2];
  size_t nSize = sizeof(dwSign);
  if (m_inStream.Read(&dwSign, &nSize))
    return O2_ERROR_READ;
  if ((nSize != sizeof(dwSign)) ||
      (dwSign[0] != OLE2_SIGNATURE) ||
      (dwSign[1] != OLE2_VERSION))
    return O2_ERROR_SIGN;

  // �������� ����� OLE2
  O2Res res = m_ole2File.Open(&m_inStream);
  if (!m_ole2File.GetOpened())
    return res;

  // ����������� ���� �����-����������
  DetectFileType();

  // ������������� ������ ����������� ���� ���������
  InitItemNameAliasList();

  // ���������� ���� � �����-����������
  m_pwszFilePath = StrDup(pwszFileName);
  if (m_pwszFilePath)
  {
    // ��������� ����� �����-����������
    m_pwszFileName = ::GetFileName(m_pwszFilePath);
  }

  // ��������� ������ ������ ��� ���� � �������� ��������
  m_pwszCurrentDirPath = (wchar_t *)malloc(DEFAULT_CURRENTDIRPATH_LEN *
                                           sizeof(wchar_t));
  if (m_pwszCurrentDirPath)
    m_cchCurrentDirPathBuf = DEFAULT_CURRENTDIRPATH_LEN;
  // ���������� ���� � �������� ��������
  UpdateCurrentDirPath();

  // ���������� ���������� � �����-����������
  UpdateFileInfo();

  return res;
}
/***************************************************************************/
/* Close - ��������                                                        */
/***************************************************************************/
void CStorageObject::Close()
{
  m_pwszFileName = NULL;
  m_nFileType = OLE2_UNKNOWN;
  m_nNumStreams = 0;
  m_nNumStorages = 0;
  m_nNumMacros = 0;
  m_nFindItemIndex = 0;
  m_dwCurParentItemID = 0;
  m_bCurParentItemIsStream = false;
  m_cchCurrentDirPathBuf = 0;
  m_nNumItemNameAliases = 0;
  // �������� ����� OLE2
  m_ole2File.Close();
  // �������� ��������� ������
  m_inStream.Close();
  // ����������� ���� � �����-����������
  if (m_pwszFilePath)
  {
    free(m_pwszFilePath);
    m_pwszFilePath = NULL;
  }
  // ����������� ���� � �������� ��������
  if (m_pwszCurrentDirPath)
  {
    free(m_pwszCurrentDirPath);
    m_pwszCurrentDirPath = NULL;
  }
  // ����������� ������ ����������� ���� ���������
  if (m_pItemNameAliases)
  {
    free(m_pItemNameAliases);
    m_pItemNameAliases = NULL;
  }
}
/***************************************************************************/
/* GetFileSize - ��������� ������� �����                                   */
/***************************************************************************/
unsigned __int64 CStorageObject::GetFileSize() const
{
  if (!m_ole2File.GetOpened())
    return 0i64;
  return (unsigned __int64)m_ole2File.GetSectorSize() *
                           (m_ole2File.GetFATEntryCount() + 1);
}
/***************************************************************************/
/* GetItemCount - ��������� ���������� ��������� � ������� ��������        */
/***************************************************************************/
size_t CStorageObject::GetItemCount(
  bool bIncludeParentItem
  ) const
{
  if (!m_ole2File.GetOpened())
    return 0;

  size_t nItemCount = 0;

  if (bIncludeParentItem)
  {
    // ������������ �������
    if (m_ole2File.GetDirEntry(m_dwCurParentItemID))
      nItemCount++;
  }

  if (m_bCurParentItemIsStream)
    return (nItemCount + 1);

  for (size_t i = 0; i < m_ole2File.GetDirEntryItemCount(); i++)
  {
    if (m_dwCurParentItemID == m_ole2File.GetDirEntryItem(i)->dwParentID)
      nItemCount++;
  }
  return nItemCount;
}
/***************************************************************************/
/* FindFirstItem - ����� ������� �������� �������� ��������                */
/***************************************************************************/
bool CStorageObject::FindFirstItem(
  bool bIncludeParentItem,
  PFindItemData pFindItemData
  )
{
  if (!m_ole2File.GetOpened())
    return false;

  m_nFindItemIndex = 0;

  if (bIncludeParentItem)
  {
    // ������������ �������
    const DIRECTORY_ENTRY *pDirEntry;
    pDirEntry = m_ole2File.GetDirEntry(m_dwCurParentItemID);
    if (pDirEntry)
    {
      // �������� ��������
      pFindItemData->dwItemAttributes = FILE_ATTRIBUTE_DIRECTORY;
      // ���� ��������
      pFindItemData->ftCreationTime.dwLowDateTime =
        pDirEntry->dwLowCreationTime;
      pFindItemData->ftCreationTime.dwHighDateTime =
        pDirEntry->dwHighCreationTime;
      // ���� ���������
      pFindItemData->ftModifiedTime.dwLowDateTime =
        pDirEntry->dwLowModifiedTime;
      pFindItemData->ftModifiedTime.dwHighDateTime =
        pDirEntry->dwHighModifiedTime;
      // ������ ��������
      pFindItemData->nItemSize = 0i64;
      // ��� ��������
      pFindItemData->pwszItemName = PARENT_ITEM_NAME;
      // ������������� ��������
      pFindItemData->dwItemID = m_dwCurParentItemID;
      return true;
    }
  }

  // ����� ���������� �������� �������� ��������
  return FindNextItem(pFindItemData);
}
/***************************************************************************/
/* FindNextItem - ����� ���������� �������� �������� ��������              */
/***************************************************************************/
bool CStorageObject::FindNextItem(
  PFindItemData pFindItemData
  )
{
  if (!m_ole2File.GetOpened())
    return false;

  if (m_bCurParentItemIsStream)
  {
    // ����� ���������� ���������� ��������
    if (m_nFindItemIndex == 0)
    {
      m_nFindItemIndex++;

      const DIRECTORY_ENTRY *pDirEntry;
      pDirEntry = m_ole2File.GetDirEntry(m_dwCurParentItemID);
      if (pDirEntry)
      {
        // ���� ��������
        pFindItemData->ftCreationTime.dwLowDateTime =
          pDirEntry->dwLowCreationTime;
        pFindItemData->ftCreationTime.dwHighDateTime =
          pDirEntry->dwHighCreationTime;
        // ���� ���������
        pFindItemData->ftModifiedTime.dwLowDateTime =
          pDirEntry->dwLowModifiedTime;
        pFindItemData->ftModifiedTime.dwHighDateTime =
          pDirEntry->dwHighModifiedTime;
      }
      else
      {
        // ���� ��������
        pFindItemData->ftCreationTime.dwLowDateTime = 0;
        pFindItemData->ftCreationTime.dwHighDateTime = 0;
        // ���� ���������
        pFindItemData->ftModifiedTime.dwLowDateTime = 0;
        pFindItemData->ftModifiedTime.dwHighDateTime = 0;
      }
      // �������� ��������
      pFindItemData->dwItemAttributes = 0;
      // ������ ��������
      pFindItemData->nItemSize = m_dummyItem.nItemSize;
      // ��� ��������
      pFindItemData->pwszItemName = m_dummyItem.wszItemName;
      // ������������� ��������
      pFindItemData->dwItemID = NOSTREAM;
      return true;
    }
  }
  else
  {
    // ����� ���������� �������� ������ ��������
    while (m_nFindItemIndex < m_ole2File.GetDirEntryItemCount())
    {
      const DirEntryItem *pDirEntryItem;
      pDirEntryItem = m_ole2File.GetDirEntryItem(m_nFindItemIndex);
      m_nFindItemIndex++;
      if (m_dwCurParentItemID == pDirEntryItem->dwParentID)
      {
        if ((pDirEntryItem->nType == DIR_ENTRY_TYPE_STORAGE) ||
            (pDirEntryItem->nType == DIR_ENTRY_TYPE_LOST))
        {
          // �������� ��������
          pFindItemData->dwItemAttributes = FILE_ATTRIBUTE_DIRECTORY;
          // ������ ��������
          pFindItemData->nItemSize = 0i64;
        }
        else
        {
          // �������� ��������
          pFindItemData->dwItemAttributes = 0;
          if (UseOpenableStreamAttr)
          {
            // ����������� ���� ������
            StreamType streamType = DetectStream(pDirEntryItem->dwID);
            if ((streamType == STREAM_VBA) ||
                (streamType == STREAM_WORDDOC) ||
                (streamType == STREAM_SUMINFO) ||
                (streamType == STREAM_DOCSUMINFO) ||
                (streamType == STREAM_NATIVEDATA))
              pFindItemData->dwItemAttributes = OPENABLE_STREAM_ATTRIBUTES;
          }
          // ������ ��������
          pFindItemData->nItemSize =
            (m_ole2File.GetMajorVersion() > 3)
              ? pDirEntryItem->pDirEntry->qwStreamSize
              : (unsigned __int64)pDirEntryItem->pDirEntry->dwLowStreamSize;
        }
        // ���� ��������
        pFindItemData->ftCreationTime.dwLowDateTime =
          pDirEntryItem->pDirEntry->dwLowCreationTime;
        pFindItemData->ftCreationTime.dwHighDateTime =
          pDirEntryItem->pDirEntry->dwHighCreationTime;
        // ���� ���������
        pFindItemData->ftModifiedTime.dwLowDateTime =
          pDirEntryItem->pDirEntry->dwLowModifiedTime;
        pFindItemData->ftModifiedTime.dwHighDateTime =
          pDirEntryItem->pDirEntry->dwHighModifiedTime;
        // ��� ��������
        // ����� ���������� ����� ��������
        pFindItemData->pwszItemName = FindItemNameAlias(pDirEntryItem->dwID);
        if (!pFindItemData->pwszItemName)
        {
          pFindItemData->pwszItemName =
            (const wchar_t *)pDirEntryItem->pDirEntry->bName;
        }
        // ������������� ��������
        pFindItemData->dwItemID = pDirEntryItem->dwID;
        return true;
      }
    }
  }

  return false;
}
/***************************************************************************/
/* SetCurrentDir - ��������� �������� ��������                             */
/***************************************************************************/
bool CStorageObject::SetCurrentDir(
  const wchar_t *pwszDirPath
  )
{
  if (!pwszDirPath || !pwszDirPath[0] ||
      !m_ole2File.GetOpened())
    return false;

  uint32_t dwDirEntryID;

  // ��������� ������ �������� �� ����
  dwDirEntryID = m_ole2File.GetDirEntryByPath(pwszDirPath,
                                              m_dwCurParentItemID,
                                              true);
  if (dwDirEntryID == NOSTREAM)
    return false;

  // ����������?
  const DIRECTORY_ENTRY *pDirEntry;
  pDirEntry = m_ole2File.GetDirEntry(dwDirEntryID);
  if (!pDirEntry ||
      ((pDirEntry->bObjectType != OBJ_TYPE_STORAGE) &&
       (pDirEntry->bObjectType != OBJ_TYPE_ROOT)))
    return false;

  // ��������� �������� ������������� �������� ��������
  SetCurParentItemID(dwDirEntryID, false);

  return true;
}
/***************************************************************************/
/* OpenStream - �������� ������                                            */
/***************************************************************************/
bool CStorageObject::OpenStream(
  unsigned long dwStreamID
  )
{
  if (!m_ole2File.GetOpened() ||
      (dwStreamID == NOSTREAM))
    return false;

  O2Res res;

  // ����������� ���� ������
  StreamType streamType = DetectStream(dwStreamID);

  size_t nItemSize = 0;

  m_dummyItem.wszItemName[0] = L'\0';

  switch (streamType)
  {
    case STREAM_VBA:
      // ����������� ������� ��������� ���� ������� VBA
      res = ExtractVBAMacroSource(&m_ole2File, dwStreamID, NULL, &nItemSize);
      break;
    case STREAM_WORDDOC:
      // ����������� ������� ������ ��������� Word
      res = ExtractWordDocumentText(&m_ole2File, dwStreamID, NULL,
                                    &nItemSize);
      break;
    case STREAM_SUMINFO:
    case STREAM_DOCSUMINFO:
      // ����������� ������� ����������
      res = ExtractMetadata(&m_ole2File, dwStreamID,
                            (streamType == STREAM_SUMINFO)
                              ? PSS_SUMMARYINFO
                              : PSS_DOCSUMMARYINFO,
                            NULL, &nItemSize);
      break;
    case STREAM_NATIVEDATA:
      // ��������� ���������� � ��������� �������
      res = GetEmbeddedObjectInfo(&m_ole2File, dwStreamID,
                                  m_dummyItem.wszItemName,
                                  countof(m_dummyItem.wszItemName),
                                  &nItemSize);
      break;
  }

  if (nItemSize == 0)
    return false;

  // ������������� ���������� ��������
  // ������ ��������
  m_dummyItem.nItemSize = nItemSize;
  // ��� � ��� ��������
  const wchar_t *pwszItemName = NULL;
  switch (streamType)
  {
    case STREAM_VBA:
      m_dummyItem.nItemType = STREAM_DATA_TYPE_VBA_MACRO_SRC;
      // ��������� ����� ����� ��������� ���� ������� VBA
      MakeVBAMacroSourceFileName(&m_ole2File, dwStreamID,
                                 m_dummyItem.wszItemName,
                                 countof(m_dummyItem.wszItemName));
      break;
    case STREAM_WORDDOC:
      m_dummyItem.nItemType = STREAM_DATA_TYPE_WORDDOC_TEXT;
      pwszItemName = WORDDOC_TEXT_FILE_NAME;
      break;
    case STREAM_SUMINFO:
    case STREAM_DOCSUMINFO:
      m_dummyItem.nItemType = (streamType == STREAM_SUMINFO)
                                ? STREAM_DATA_TYPE_SUMINFO_TEXT
                                : STREAM_DATA_TYPE_DOCSUMINFO_TEXT;
      pwszItemName = INFO_TEXT_FILE_NAME;
      break;
    case STREAM_NATIVEDATA:
      m_dummyItem.nItemType = STREAM_DATA_TYPE_NATIVE_DATA;
      break;
  }

  if (pwszItemName)
  {
    size_t cchItemName = ::lstrlen(pwszItemName);
    if (cchItemName >= countof(m_dummyItem.wszItemName))
      cchItemName = countof(m_dummyItem.wszItemName) - 1;
    memcpy(m_dummyItem.wszItemName, pwszItemName,
           cchItemName * sizeof(wchar_t));
    m_dummyItem.wszItemName[cchItemName] = L'\0';
  }

  // ��������� �������� ������������� �������� ��������
  SetCurParentItemID(dwStreamID, true);

  return true;
}
/***************************************************************************/
/* DetectStream - ����������� ���� ������                                  */
/***************************************************************************/
StreamType CStorageObject::DetectStream(
  unsigned long dwStreamID
  ) const
{
  // ��������, �������� ������ �������� �������
  const DIRECTORY_ENTRY *pDirEntry;
  pDirEntry = m_ole2File.GetDirEntry(dwStreamID);
  if (!pDirEntry ||
      (pDirEntry->bObjectType != OBJ_TYPE_STREAM))
    return STREAM_UNKNOWN;

  // �������������� ������ VBA
  if (DetectVBAStream(&m_ole2File, dwStreamID))
    return STREAM_VBA;

  size_t cchStreamType = pDirEntry->wSizeOfName >> 1;
  if (cchStreamType <= 1)
    return STREAM_UNKNOWN;
  cchStreamType--;

  for (size_t i = 0; i < countof(OPENABLE_STREAM_NAMES); i++)
  {
    if (!StrCmpNIW((const wchar_t *)pDirEntry->bName,
                   OPENABLE_STREAM_NAMES[i],
                   cchStreamType) &&
        (OPENABLE_STREAM_NAMES[i][cchStreamType] == L'\0'))
      return (StreamType)(STREAM_WORDDOC + i);
  }
  return STREAM_UNKNOWN;
}
/***************************************************************************/
/* SetFileItemTime - ��������� ������� ����� �� ������� ������ ��������    */
/***************************************************************************/
bool SetFileItemTime(
  const wchar_t *pwszFileName,
  const DIRECTORY_ENTRY *pDirEntry
  )
{
  const FILETIME *pftCreationTime = NULL;
  const FILETIME *pftLastWriteTime = NULL;

  if ((pDirEntry->dwLowCreationTime != 0) ||
      (pDirEntry->dwHighCreationTime != 0))
    pftCreationTime = (FILETIME *)pDirEntry->bCreationTime;
  if ((pDirEntry->dwLowModifiedTime != 0) ||
      (pDirEntry->dwHighModifiedTime != 0))
    pftLastWriteTime = (FILETIME *)pDirEntry->bModifiedTime;

  // ��������� ������� �����
  if (!pftCreationTime && !pftLastWriteTime)
    return true;

  HANDLE hFile;

  // �������� �����
  hFile = ::CreateFile(pwszFileName, FILE_WRITE_ATTRIBUTES,
                       FILE_SHARE_READ | FILE_SHARE_WRITE,
                       NULL, OPEN_EXISTING,
                       FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_NO_BUFFERING |
                       FILE_FLAG_SEQUENTIAL_SCAN,
                       NULL);
  if (hFile == INVALID_HANDLE_VALUE)
    return false;

  // ��������� ������� �����
  bool bSuccess = (0 != ::SetFileTime(hFile, pftCreationTime, NULL,
                                      pftLastWriteTime));

  ::CloseHandle(hFile);

  return bSuccess;
}
/***************************************************************************/
/* ExtractItem - ���������� ��������                                       */
/***************************************************************************/
bool CStorageObject::ExtractItem(
  unsigned long dwItemID,
  const wchar_t *pwszItemName,
  const wchar_t *pwszDestPath
  ) const
{
  if (!m_ole2File.GetOpened())
    return false;

  const DIRECTORY_ENTRY *pDirEntry;

  if (dwItemID != NOSTREAM)
  {
    // ������ ��������
    pDirEntry = m_ole2File.GetDirEntry(dwItemID);
  }
  else
  {
    // ��������� �������
    if (!m_bCurParentItemIsStream)
      return false;
    pDirEntry = m_ole2File.GetDirEntry(m_dwCurParentItemID);
  }

  if (!pDirEntry)
    return false;

  // �������� �������� ���������
  if (!ForceDirectories(pwszDestPath))
    return false;

  // ��������� ����� �����
  wchar_t *pwszFileName = AllocAndCombinePath(pwszDestPath, pwszItemName);
  if (!pwszFileName)
    return false;

  bool bSuccess = false;

  if (dwItemID != NOSTREAM)
  {
    // ������ ��������
    if ((pDirEntry->bObjectType == OBJ_TYPE_STORAGE) ||
        (pDirEntry->bObjectType == OBJ_TYPE_ROOT))
    {
      // �������� �������� ���������
      if (ForceDirectories(pwszFileName))
      {
        // ����������� ���������� ��������� ��������
        if (ExtractStorageItems(dwItemID, pwszFileName))
          bSuccess = true;
      }
    }
    else
    {
      // ���������� ������
      if (ExtractStream(dwItemID, pwszFileName))
        bSuccess = true;
    }
  }
  else
  {
    // ��������� �������
    // ���������� ������ ������
    if (ExtractStreamData(m_dwCurParentItemID, m_dummyItem.nItemType,
                          pwszFileName))
      bSuccess = true;
  }

  if (bSuccess)
  {
    // ��������� ������� ����� �� ������� ������ ��������
    SetFileItemTime(pwszFileName, pDirEntry);
  }

  free(pwszFileName);

  return bSuccess;
}
/***************************************************************************/
/* ExtractStorageItems - ���������� ��������� ��������                     */
/***************************************************************************/
bool CStorageObject::ExtractStorageItems(
  unsigned long dwStorageID,
  const wchar_t *pwszDestPath
  ) const
{
  bool bError = false;
  for (size_t i = 0; i < m_ole2File.GetDirEntryItemCount(); i++)
  {
    const DirEntryItem *pDirEntryItem = m_ole2File.GetDirEntryItem(i);
    if (dwStorageID == pDirEntryItem->dwParentID)
    {
      // ����� ���������� ����� ��������
      const wchar_t *pwszItemName = FindItemNameAlias(pDirEntryItem->dwID);
      if (!pwszItemName)
        pwszItemName = (const wchar_t *)pDirEntryItem->pDirEntry->bName;
      // ���������� ��������
      if (!ExtractItem(pDirEntryItem->dwID, pwszItemName, pwszDestPath))
        bError = true;
    }
  }
  return !bError;
}
/***************************************************************************/
/* ExtractStream - ���������� ������                                       */
/***************************************************************************/
bool CStorageObject::ExtractStream(
  unsigned long dwStreamID,
  const wchar_t *pwszFileName
  ) const
{
  bool bSuccess = false;

  // �������� ������� ��������� ������ ��� ����������������� ������ ������
  CFileSeqOutStream outStream;

  // �������� �����
  if (!outStream.Open(pwszFileName))
  {
    // ���������� ������
    if (O2_OK == m_ole2File.ExtractStream(dwStreamID, &outStream))
      bSuccess = true;
  }

  return bSuccess;
}
/***************************************************************************/
/* ExtractStreamData - ���������� ������ ������                            */
/***************************************************************************/
bool CStorageObject::ExtractStreamData(
  unsigned long dwStreamID,
  int nStreamDataType,
  const wchar_t *pwszFileName
  ) const
{
  O2Res res;

  size_t nProcessedSize = 0;

  // �������� ������� ��������� ������ ��� ����������������� ������ ������
  CFileSeqOutStream outStream;

  // �������� �����
  if (!outStream.Open(pwszFileName))
  {
    switch (nStreamDataType)
    {
      case STREAM_DATA_TYPE_VBA_MACRO_SRC:
        // ���������� ��������� ���� ������� VBA
        res = ExtractVBAMacroSource(&m_ole2File, dwStreamID, &outStream,
                                    &nProcessedSize);
        break;
      case STREAM_DATA_TYPE_WORDDOC_TEXT:
        // ���������� ������ ��������� Word
        res = ExtractWordDocumentText(&m_ole2File, dwStreamID, &outStream,
                                      &nProcessedSize);
        break;
      case STREAM_DATA_TYPE_SUMINFO_TEXT:
      case STREAM_DATA_TYPE_DOCSUMINFO_TEXT:
        {
          // ���������� ����������
          PropertySetStreamType streamType;
          if (m_dummyItem.nItemType == STREAM_DATA_TYPE_SUMINFO_TEXT)
            streamType = PSS_SUMMARYINFO;
          else
            streamType = PSS_DOCSUMMARYINFO;
          res = ExtractMetadata(&m_ole2File, dwStreamID, streamType,
                                &outStream, &nProcessedSize);
        }
        break;
      case STREAM_DATA_TYPE_NATIVE_DATA:
        // ���������� ���������� ������� �� ������ "\1Ole10Native"
        ExtractEmbeddedObject(&m_ole2File, dwStreamID, &outStream,
                              &nProcessedSize);
        break;
    }
  }

  return (nProcessedSize != 0);
}
/***************************************************************************/
/* SetCurParentItemID - ��������� �������� ������������� �������� �������� */
/***************************************************************************/
void CStorageObject::SetCurParentItemID(
  unsigned long dwDirEntryID,
  bool bIsStream
  )
{
  if (dwDirEntryID != m_dwCurParentItemID)
  {
    m_dwCurParentItemID = dwDirEntryID;
    m_bCurParentItemIsStream = bIsStream;
    m_nFindItemIndex = 0;
    // ���������� ���� � �������� ��������
    UpdateCurrentDirPath();
  }
}
/***************************************************************************/
/* UpdateCurrentDirPath - ���������� ���� � �������� ��������              */
/***************************************************************************/
void CStorageObject::UpdateCurrentDirPath()
{
  // ��������� ������� ���� ������ ��������
  size_t cch = m_ole2File.GetDirEntryFullPath(m_dwCurParentItemID,
                                              m_pwszCurrentDirPath,
                                              m_cchCurrentDirPathBuf,
                                              true);
  if (cch == 0)
  {
    if (m_pwszCurrentDirPath)
      m_pwszCurrentDirPath[0] = L'\0';
  }
  else if (cch >= m_cchCurrentDirPathBuf)
  {
    if (m_pwszCurrentDirPath)
      free(m_pwszCurrentDirPath);
    m_pwszCurrentDirPath = (wchar_t *)malloc(cch * sizeof(wchar_t));
    if (!m_pwszCurrentDirPath)
    {
      m_cchCurrentDirPathBuf = 0;
      return;
    }
    m_cchCurrentDirPathBuf = cch;
    // ��������� ������� ���� ������ ��������
    m_ole2File.GetDirEntryFullPath(m_dwCurParentItemID, m_pwszCurrentDirPath,
                                   cch, true);
  }
}
/***************************************************************************/
/* DetectFileType - ����������� ���� �����-����������                      */
/***************************************************************************/
void CStorageObject::DetectFileType()
{
  // ������������ �������� ������� ��������� ����� ������ OLE2
  const wchar_t* const MAIN_STREAM_NAMES[] =
  {
    L"WordDocument",
    L"Workbook",
    L"PowerPoint Document",
    L"VisioDocument",
    L"Book",
    L"DocumentText"
  };

  // �������������� �����-���������� MSI
  if (DetectMsi(&m_ole2File))
  {
    m_nFileType = OLE2_MSI;
    return;
  }

  // ����� � �������� �������� �������� ������� ��������� ����� ������ OLE2
  m_nFileType = OLE2_UNKNOWN;
  for (size_t i = 0; i < countof(MAIN_STREAM_NAMES); i++)
  {
    if (NOSTREAM != m_ole2File.FindDirEntry(0, OBJ_TYPE_STREAM,
                                            MAIN_STREAM_NAMES[i],
                                            true))
    {
      m_nFileType = (OLE2FileType)(OLE2_MSWORD + i);
      break;
    }
  }
}
/***************************************************************************/
/* UpdateFileInfo - ���������� ���������� � �����-����������               */
/***************************************************************************/
void CStorageObject::UpdateFileInfo()
{
  // ����������� ���������� �������, ��������� � ��������
  m_nNumStreams = 0;
  m_nNumStorages = 0;
  m_nNumMacros = 0;
  for (size_t i = 1; i < m_ole2File.GetDirEntryItemCount(); i++)
  {
    const DirEntryItem *pDirEntryItem;
    pDirEntryItem = m_ole2File.GetDirEntryItem(i);
    if (pDirEntryItem->nType == DIR_ENTRY_TYPE_STORAGE)
    {
      m_nNumStorages++;
    }
    else if (pDirEntryItem->nType == DIR_ENTRY_TYPE_STREAM)
    {
      m_nNumStreams++;
      // �������������� ������ VBA
      if (DetectVBAStream(&m_ole2File, pDirEntryItem->dwID))
      {
        // ����������� ������� ��������� ���� ������� VBA
        size_t cbMacroSrc;
        ExtractVBAMacroSource(&m_ole2File, pDirEntryItem->dwID, NULL,
                              &cbMacroSrc);
        if (cbMacroSrc != 0)
          m_nNumMacros++;
      }
    }
  }
}
/***************************************************************************/
/* InitItemNameAliasList - ������������� ������ ����������� ���� ��������� */
/***************************************************************************/
void CStorageObject::InitItemNameAliasList()
{
  const DirEntryItem *pDirEntryItem;

  size_t nNumAliases = 0;

  // ����������� ���������� ����������� � ������
  for (size_t i = 1; i < m_ole2File.GetDirEntryItemCount(); i++)
  {
    pDirEntryItem = m_ole2File.GetDirEntryItem(i);
    // �������� ������������� ������������� ����� ������ �������� MSI ���
    // ������� � ����� ������ �������� ������������ ��������
    if (((m_nFileType == OLE2_MSI) &&
         DecodeMsiDirEntryName(pDirEntryItem->pDirEntry->bName, NULL, 0)) ||
        !IsDirEntryNameValid(pDirEntryItem->pDirEntry->bName))
    {
      nNumAliases++;
    }
  }

  if (nNumAliases == 0)
    return;

  m_pItemNameAliases = (PDirEntryNameAlias)malloc(nNumAliases *
                                                  sizeof(DirEntryNameAlias));
  if (!m_pItemNameAliases)
    return;

  size_t nAliasCount = 0;

  PDirEntryNameAlias pAlias = m_pItemNameAliases;
  for (size_t i = 1;
       (i < m_ole2File.GetDirEntryItemCount()) &&
       (nAliasCount < nNumAliases);
       i++)
  {
    pDirEntryItem = m_ole2File.GetDirEntryItem(i);
    size_t cchAlias = 0;
    if (m_nFileType == OLE2_MSI)
    {
      // ������������� ����� ������ �������� MSI
      cchAlias = DecodeMsiDirEntryName(pDirEntryItem->pDirEntry->bName,
                                       pAlias->wszNameAlias,
                                       sizeof(pAlias->wszNameAlias));
    }
    if (cchAlias == 0)
    {
      // ��������, �������� �� ��� ������ �������� ������������ �������
      if (!IsDirEntryNameValid(pDirEntryItem->pDirEntry->bName))
      {
        // �������������� ����� ������ �������� � ��� �����
        cchAlias = DirEntryNameToFileName(pDirEntryItem->pDirEntry->bName,
                                          pAlias->wszNameAlias,
                                          sizeof(pAlias->wszNameAlias));
      }
    }
    if ((cchAlias == 0) || (cchAlias >= sizeof(pAlias->wszNameAlias)))
      continue;
    // �������� ���������� ������ �������� � ������, ���������������
    // ����������
    if (NOSTREAM == m_ole2File.FindDirEntry(pDirEntryItem->dwParentID, 0,
                                            pAlias->wszNameAlias, false))
    {
      pAlias->dwDirEntryID = pDirEntryItem->dwID;
      pAlias++;
      nAliasCount++;
    }
  }

  m_nNumItemNameAliases = nAliasCount;

  // ��������� ������ ����������� ���� ������� ��������
  m_ole2File.SetDirEntryNameAliases(m_pItemNameAliases,
                                    m_nNumItemNameAliases);
}
/***************************************************************************/
/* FindItemNameAlias - ����� ���������� ����� ��������                     */
/***************************************************************************/
const wchar_t *CStorageObject::FindItemNameAlias(
  unsigned long dwItemID
  ) const
{
  for (size_t i = 0; i < m_nNumItemNameAliases; i++)
  {
    if (dwItemID == m_pItemNameAliases[i].dwDirEntryID)
      return m_pItemNameAliases[i].wszNameAlias;
  }
  return NULL;
}
//---------------------------------------------------------------------------
