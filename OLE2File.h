//---------------------------------------------------------------------------
#ifndef __OLE2FILE_H__
#define __OLE2FILE_H__
//---------------------------------------------------------------------------
#include "OLE2Type.h"
#include "OLE2CF.h"
#include "Streams.h"
//---------------------------------------------------------------------------
/***************************************************************************/
/*                                                                         */
/*                      COLE2File - ����� ����� OLE2                       */
/*                                                                         */
/***************************************************************************/

// ������������� ���������� �������� ��� "����������" �������
#define LOSTSTORAGEID  0xFFFFFFFE

// ��� ������ ��������
typedef enum
{
  DIR_ENTRY_TYPE_UNKNOWN,
  DIR_ENTRY_TYPE_STORAGE,
  DIR_ENTRY_TYPE_STREAM,
  DIR_ENTRY_TYPE_LOST
} DirEntryType;

// ������� ������ ��������
typedef struct _DirEntryItem
{
  int                    nType;
  uint32_t               dwID;
  uint32_t               dwParentID;
  uint32_t               dwFirstChildID;
  const DIRECTORY_ENTRY *pDirEntry;
} DirEntryItem, *PDirEntryItem;

// ������������ ����� ���������� ����� ������ ��������
#define MAX_DIR_ENTRY_NAME_ALIAS_LEN  DIRECTORY_ENTRY_SIZE

// ��������� ����� ������ ��������
typedef struct _DirEntryNameAlias
{
  uint32_t  dwDirEntryID;
  wchar_t   wszNameAlias[MAX_DIR_ENTRY_NAME_ALIAS_LEN];
} DirEntryNameAlias, *PDirEntryNameAlias;

class COLE2File
{
public:
  // ��������
  O2Res Open(
    CInStream *pInStream
    );
  // ��������
  void Close();
  // ��������� ����� �������� �����
  inline bool GetOpened() const
    { return m_bOpened; }
  // ��������� ������ ��� ����� ������
  inline CInStream *GetInStream() const
    { return m_pInStream; }
  // ��������� c������� ������ ������
  inline int GetMajorVersion() const
    { return m_nMajorVersion; }
  // ��������� �������� ������ ������
  inline int GetMinorVersion() const
    { return m_nMinorVersion; }
  // ��������� ��������� �����
  inline const COMPOUND_FILE_HEADER *GetFileHeader() const
    { return &m_fileHeader; }
  // ��������� ������� �������
  inline size_t GetSectorSize() const
    { return m_nSectorSize; }
  // ��������� ������� ������� mini FAT
  inline size_t GetMiniSectorSize() const
    { return m_nMiniSectorSize; }
  // ��������� FAT
  inline const uint32_t *GetFAT() const
    { return m_pFAT; }
  // ��������� ���������� ������� FAT
  inline size_t GetFATEntryCount() const
    { return m_nNumFATEntries; }
  // ��������� mini FAT
  inline const uint32_t *GetMiniFAT() const
    { return m_pMiniFAT; }
  // ��������� ���������� ������� mini FAT
  inline size_t GetMiniFATEntryCount() const
    { return m_nNumMiniFATEntries; }
  // ��������� ��������
  inline const DIRECTORY_ENTRY *GetDirectory() const
    { return m_pDirEntries; }
  // ��������� ���������� ������� ��������
  inline size_t GetDirEntryCount() const
    { return m_nNumDirEntries; }
  // ��������� ������ ��������
  const DIRECTORY_ENTRY *GetDirEntry(
    uint32_t dwID
    ) const;
  // ��������� �������� ������ ��������
  const DirEntryItem *GetDirEntryItem(
    size_t nIndex
    ) const;
  // ��������� ���������� ��������� ������� ��������
  inline size_t GetDirEntryItemCount() const
    { return m_nNumDirEntryItems; }
  // ��������� ����� ������� "����������" ������� ��������
  inline bool GetSomeDirEntriesLost() const
    { return m_bSomeDirEntriesLost; }
  // ��������� ������ ����������� ���� ������� ��������
  void SetDirEntryNameAliases(
    const DirEntryNameAlias *pDirEntryNameAliases,
    size_t nNumDirEntryNameAliases
    );
  // ����� ���������� ����� ������ ��������
  const wchar_t *FindDirEntryNameAlias(
    uint32_t dwDirEntryID
    ) const;
  // ��������� ������ �������� �� ����
  uint32_t GetDirEntryByPath(
    const wchar_t *pwszPath,
    uint32_t dwBaseStorageID,
    bool bUseAliases
    ) const;
  // ����� ������ ��������
  uint32_t FindDirEntry(
    uint32_t dwStorageID,
    uint8_t bObjectType,
    const wchar_t *pwszName,
    bool bUseAliases
    ) const;
  // ��������� �������������� ������������ ������ ��������
  uint32_t GetParentDirEntryID(
    uint32_t dwDirEntryID
    ) const;
  // ��������� ������� ���� ������ ��������
  size_t GetDirEntryFullPath(
    uint32_t dwDirEntryID,
    wchar_t *pBuffer,
    size_t nBufferSize,
    bool bUseAliases
    ) const;
  // ���������� ����������� ������
  O2Res ExtractStreamData(
    uint32_t dwStreamID,
    uint64_t nOffset,
    size_t nSize,
    void *pBuffer,
    size_t *pnProcessedSize = NULL
    ) const;
  // ���������� ����������� ������
  O2Res ExtractStreamData(
    uint32_t dwStreamID,
    uint64_t nOffset,
    uint64_t nSize,
    CSeqOutStream *pOutStream,
    uint64_t *pnProcessedSize = NULL
    ) const;
  // ���������� ������
  O2Res ExtractStream(
    uint32_t dwStreamID,
    CSeqOutStream *pOutStream,
    uint64_t *pnProcessedSize = NULL
    ) const;

  // ����������� ������
  COLE2File();
  // ���������� ������
  virtual ~COLE2File();

private:
  // ����������� � ���������� �� ���������
  COLE2File(const COLE2File&);
  void operator=(const COLE2File&);

protected:
  // �������� ��������� �����
  virtual bool CheckFileHeader() const;

private:
  bool                     m_bOpened;
  CInStream               *m_pInStream;
  int                      m_nMajorVersion;
  int                      m_nMinorVersion;
  size_t                   m_nEffectiveNumSectors;
  size_t                   m_nEffectiveNumMiniSectors;
  uint32_t                 m_dwRootEntryStartSector;
  size_t                   m_nSectorSize;
  size_t                   m_nMiniSectorSize;
  size_t                   m_nMiniStreamCutoffSize;
  uint32_t                *m_pFAT;
  size_t                   m_nNumFATEntries;
  uint32_t                *m_pMiniFAT;
  size_t                   m_nNumMiniFATEntries;
  PDIRECTORY_ENTRY         m_pDirEntries;
  size_t                   m_nNumDirEntries;
  bool                     m_bSomeDirEntriesLost;
  uint64_t                 m_nFileSize;
  COMPOUND_FILE_HEADER     m_fileHeader;
  PDirEntryItem            m_pDirEntryItems;
  size_t                   m_nNumDirEntryItems;
  PDIRECTORY_ENTRY         m_pLostStorageDirEntry;
  const DirEntryNameAlias *m_pDirEntryNameAliases;
  size_t                   m_nNumDirEntryNameAliases;

  // �������������
  O2Res Init(
    CInStream *pInStream
    );
  // ����� ������ ��������
  uint32_t FindDirEntry(
    uint32_t dwStorageID,
    uint8_t bObjectType,
    const wchar_t *pwchName,
    size_t cchName,
    bool bUseAliases
    ) const;
  // ����������� ������� ��������� ������� ���� ������ ��������
  size_t RecurseGetDirEntryFullPath(
    uint32_t dwDirEntryID,
    wchar_t *pBuffer,
    size_t nBufferSize,
    bool bUseAliases
    ) const;
  // ���������� ����������� ������
  O2Res ExtractStreamData(
    uint32_t dwStreamStartSector,
    uint64_t nStreamSize,
    uint64_t nOffset,
    size_t nSize,
    void *pBuffer,
    size_t *pnProcessedSize
    ) const;
  // ���������� ����������� ������
  O2Res ExtractStreamData(
    uint32_t dwStreamStartSector,
    uint64_t nStreamSize,
    uint64_t nOffset,
    uint64_t nSize,
    CSeqOutStream *pOutStream,
    uint64_t *pnProcessedSize
    ) const;
  // ��������� ������ � ������ FAT
  O2Res AllocAndReadFAT();
  // ��������� ������ � ������ mini FAT
  O2Res AllocAndReadMiniFAT();
  // ��������� ������ � ������ ��������
  O2Res AllocAndReadDirectory();
  // �������� ������ ��������� ������� ��������
  O2Res CreateDirEntryItemList();
  // ����������� ������������ ������� ��������
  void EnumDirEntries(
    uint32_t dwDirEntryID,
    uint32_t dwParentDirEntryID
    );
  // ������ ������ mini-������� � �����
  O2Res ReadMiniSectorData(
    uint32_t dwMiniSectorNum,
    size_t nOffset,
    size_t nBytesToRead,
    void *pBuffer,
    size_t *pnBytesRead
    ) const;
  // ������ ������� � �����
  O2Res ReadSector(
    uint32_t dwSectorNum,
    void *pBuffer,
    size_t *pnBytesRead
    ) const;
  // ������ ������ ������� � �����
  O2Res ReadSectorData(
    uint32_t dwSectorNum,
    size_t nOffset,
    size_t nBytesToRead,
    void *pBuffer,
    size_t *pnBytesRead
    ) const;
  // ��������� ������ ���������� mini-�������
  bool GetNextMiniSector(
    uint32_t dwMiniSectorNum,
    size_t nIndex,
    uint32_t *pdwMiniSectorNum
    ) const;
  // ��������� ������ ���������� �������
  bool GetNextSector(
    uint32_t dwSectorNum,
    size_t nIndex,
    uint32_t *pdwSectorNum
    ) const;
  // �������� ������������ ������ ������� mini FAT
  bool IsMiniSectorNumValid(
    uint32_t dwMiniSectorNum
    ) const;
  // �������� ������������ ������ �������
  bool IsSectorNumValid(
    uint32_t dwSectorNum
    ) const;
  // ��������� �������� �������
  uint64_t GetSectorOffset(
    uint32_t dwSectorNum
    ) const;
};
//---------------------------------------------------------------------------
#endif  // __OLE2FILE_H__
