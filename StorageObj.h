//---------------------------------------------------------------------------
#ifndef __STORAGEOBJ_H__
#define __STORAGEOBJ_H__
//---------------------------------------------------------------------------
#include <stddef.h>
#include "Streams.h"
#include "FileStreams.h"
#include "OLE2File.h"
//---------------------------------------------------------------------------
// �������� ��� ����������� �������
#define OPENABLE_STREAM_ATTRIBUTES  \
  (FILE_ATTRIBUTE_COMPRESSED | FILE_ATTRIBUTE_REPARSE_POINT)
//---------------------------------------------------------------------------
/***************************************************************************/
/*                                                                         */
/*                 CStorageObject - ����� �����-����������                 */
/*                                                                         */
/***************************************************************************/

// ��� ����� OLE2
typedef enum
{
  OLE2_UNKNOWN = 0,
  OLE2_MSWORD,
  OLE2_MSEXCEL,
  OLE2_MSPPOINT,
  OLE2_MSVISIO,
  OLE2_MSEXCEL5,
  OLE2_ICHITARO,
  OLE2_MSI
} OLE2FileType;

// ��� ������
typedef enum
{
  STREAM_UNKNOWN = 0,
  STREAM_VBA,
  STREAM_WORDDOC,
  STREAM_SUMINFO,
  STREAM_DOCSUMINFO,
  STREAM_NATIVEDATA
} StreamType;

// ���������� � ��������� ��������
typedef struct _FindItemData
{
  unsigned long     dwItemAttributes;
  FILETIME          ftCreationTime;
  FILETIME          ftModifiedTime;
  unsigned __int64  nItemSize;
  const wchar_t    *pwszItemName;
  unsigned long     dwItemID;
} FindItemData, *PFindItemData;

// ��� ������ ������
typedef enum
{
  STREAM_DATA_TYPE_UNKNOWN = 0,
  STREAM_DATA_TYPE_VBA_MACRO_SRC,
  STREAM_DATA_TYPE_WORDDOC_TEXT,
  STREAM_DATA_TYPE_SUMINFO_TEXT,
  STREAM_DATA_TYPE_DOCSUMINFO_TEXT,
  STREAM_DATA_TYPE_NATIVE_DATA
} StreamDataType;

// ������������ ����� ����� ���������� ��������
#define MAX_DUMMY_ITEM_NAME_LEN  MAX_PATH

// ��������� �������
typedef struct _DummyItem
{
  int      nItemType;
  size_t   nItemSize;
  wchar_t  wszItemName[MAX_DUMMY_ITEM_NAME_LEN];
} DummyItem, *PDummyItem;

class CStorageObject
{
public:
  // ���� ������������� ��������� ��� ����������� �������
  bool UseOpenableStreamAttr;

  // ��������
  int Open(
    const wchar_t *pwszFileName
    );
  // ��������
  void Close();
  // ��������� ����� �������� �����-����������
  inline bool GetOpened() const
    { return m_ole2File.GetOpened(); }
  // ��������� ���� � �����-����������
  inline const wchar_t *GetFilePath() const
    { return m_pwszFilePath; }
  // ��������� ����� �����-����������
  inline const wchar_t *GetFileName() const
    { return m_pwszFileName; }
  // ��������� c������� ������ ������
  inline int GetMajorVersion() const
    { return m_ole2File.GetMajorVersion(); }
  // ��������� �������� ������ ������
  inline int GetMinorVersion() const
    { return m_ole2File.GetMinorVersion(); }
  // ��������� ���� �����-����������
  inline int GetFileType() const
    { return m_nFileType; }
  // ��������� ���������� �������
  inline size_t GetStreamCount() const
    { return m_nNumStreams; }
  // ��������� ���������� ���������
  inline size_t GetStorageCount() const
    { return m_nNumStorages; }
  // ��������� ���������� ��������
  inline size_t GetMacroCount() const
    { return m_nNumMacros; }
  // ��������� ������� �����
  unsigned __int64 GetFileSize() const;
  // ��������� ���������� ��������� � ������� ��������
  size_t GetItemCount(
    bool bIncludeParentItem
    ) const;
  // ����� ������� �������� �������� ��������
  bool FindFirstItem(
    bool bIncludeParentItem,
    PFindItemData pFindItemData
    );
  // ����� ���������� �������� �������� ��������
  bool FindNextItem(
    PFindItemData pFindItemData
    );
  // ��������� �������� ��������
  bool SetCurrentDir(
    const wchar_t *pwszDirPath
    );
  // ��������� ���� � �������� ��������
  inline const wchar_t *GetCurrentDirPath() const
    { return m_pwszCurrentDirPath; }
  // �������� ������
  bool OpenStream(
    unsigned long dwStreamID
    );
  // ���������� ��������
  bool ExtractItem(
    unsigned long dwItemID,
    const wchar_t *pwszItemName,
    const wchar_t *pwszDestPath
    ) const;

  // ����������� ������
  CStorageObject();
  // ���������� ������
  ~CStorageObject();

private:
  // ����������� � ���������� �� ���������
  CStorageObject(const CStorageObject&);
  void operator=(const CStorageObject&);

private:
  wchar_t            *m_pwszFilePath;
  const wchar_t      *m_pwszFileName;
  int                 m_nFileType;
  size_t              m_nNumStreams;
  size_t              m_nNumStorages;
  size_t              m_nNumMacros;
  size_t              m_nFindItemIndex;
  unsigned long       m_dwCurParentItemID;
  bool                m_bCurParentItemIsStream;
  size_t              m_cchCurrentDirPathBuf;
  wchar_t            *m_pwszCurrentDirPath;
  DummyItem           m_dummyItem;
  PDirEntryNameAlias  m_pItemNameAliases;
  size_t              m_nNumItemNameAliases;
  CFileInStream       m_inStream;
  COLE2File           m_ole2File;

  // �������������
  int Init(
    const wchar_t *pwszFileName
    );
  // ����������� ���� ������
  StreamType DetectStream(
    unsigned long dwStreamID
    ) const;
  // ���������� ��������� ��������
  bool ExtractStorageItems(
    unsigned long dwStorageID,
    const wchar_t *pwszDestPath
    ) const;
  // ���������� ������
  bool ExtractStream(
    unsigned long dwStreamID,
    const wchar_t *pwszFileName
    ) const;
  // ���������� ������ ������
  bool ExtractStreamData(
    unsigned long dwStreamID,
    int nStreamDataType,
    const wchar_t *pwszFileName
    ) const;
  // ��������� �������� ������������� �������� ��������
  void SetCurParentItemID(
    unsigned long dwItemID,
    bool bIsStream
    );
  // ���������� ���� � �������� ��������
  void UpdateCurrentDirPath();
  // ����������� ���� �����-����������
  void DetectFileType();
  // ���������� ���������� � �����-����������
  void UpdateFileInfo();
  // ������������� ������ ����������� ���� ���������
  void InitItemNameAliasList();
  // ����� ���������� ����� ��������
  const wchar_t *FindItemNameAlias(
    unsigned long dwItemID
    ) const;
};
//---------------------------------------------------------------------------
// �������������� ���� ����� OLE2 � ������
const wchar_t *OLE2FileTypeToString(
  int nFileType
  );
//---------------------------------------------------------------------------
#endif  // __STORAGEOBJ_H__
