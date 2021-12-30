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
/*                      COLE2File - класс файла OLE2                       */
/*                                                                         */
/***************************************************************************/

// Идентификатор фиктивного каталога для "потерянных" записей
#define LOSTSTORAGEID  0xFFFFFFFE

// Тип записи каталога
typedef enum
{
  DIR_ENTRY_TYPE_UNKNOWN,
  DIR_ENTRY_TYPE_STORAGE,
  DIR_ENTRY_TYPE_STREAM,
  DIR_ENTRY_TYPE_LOST
} DirEntryType;

// Элемент записи каталога
typedef struct _DirEntryItem
{
  int                    nType;
  uint32_t               dwID;
  uint32_t               dwParentID;
  uint32_t               dwFirstChildID;
  const DIRECTORY_ENTRY *pDirEntry;
} DirEntryItem, *PDirEntryItem;

// Максимальная длина псевдонима имени записи каталога
#define MAX_DIR_ENTRY_NAME_ALIAS_LEN  DIRECTORY_ENTRY_SIZE

// Псевдоним имени записи каталога
typedef struct _DirEntryNameAlias
{
  uint32_t  dwDirEntryID;
  wchar_t   wszNameAlias[MAX_DIR_ENTRY_NAME_ALIAS_LEN];
} DirEntryNameAlias, *PDirEntryNameAlias;

class COLE2File
{
public:
  // Открытие
  O2Res Open(
    CInStream *pInStream
    );
  // Закрытие
  void Close();
  // Получение флага открытия файла
  inline bool GetOpened() const
    { return m_bOpened; }
  // Получение потока для ввода данных
  inline CInStream *GetInStream() const
    { return m_pInStream; }
  // Получение cтаршего номера версии
  inline int GetMajorVersion() const
    { return m_nMajorVersion; }
  // Получение младшего номера версии
  inline int GetMinorVersion() const
    { return m_nMinorVersion; }
  // Получение заголовка файла
  inline const COMPOUND_FILE_HEADER *GetFileHeader() const
    { return &m_fileHeader; }
  // Получение размера сектора
  inline size_t GetSectorSize() const
    { return m_nSectorSize; }
  // Получение размера сектора mini FAT
  inline size_t GetMiniSectorSize() const
    { return m_nMiniSectorSize; }
  // Получение FAT
  inline const uint32_t *GetFAT() const
    { return m_pFAT; }
  // Получение количества записей FAT
  inline size_t GetFATEntryCount() const
    { return m_nNumFATEntries; }
  // Получение mini FAT
  inline const uint32_t *GetMiniFAT() const
    { return m_pMiniFAT; }
  // Получение количества записей mini FAT
  inline size_t GetMiniFATEntryCount() const
    { return m_nNumMiniFATEntries; }
  // Получение каталога
  inline const DIRECTORY_ENTRY *GetDirectory() const
    { return m_pDirEntries; }
  // Получение количества записей каталога
  inline size_t GetDirEntryCount() const
    { return m_nNumDirEntries; }
  // Получение записи каталога
  const DIRECTORY_ENTRY *GetDirEntry(
    uint32_t dwID
    ) const;
  // Получение элемента записи каталога
  const DirEntryItem *GetDirEntryItem(
    size_t nIndex
    ) const;
  // Получение количества элементов записей каталога
  inline size_t GetDirEntryItemCount() const
    { return m_nNumDirEntryItems; }
  // Получение флага наличия "потерянных" записей каталога
  inline bool GetSomeDirEntriesLost() const
    { return m_bSomeDirEntriesLost; }
  // Установка списка псевдонимов имен записей каталога
  void SetDirEntryNameAliases(
    const DirEntryNameAlias *pDirEntryNameAliases,
    size_t nNumDirEntryNameAliases
    );
  // Поиск псевдонима имени записи каталога
  const wchar_t *FindDirEntryNameAlias(
    uint32_t dwDirEntryID
    ) const;
  // Получение записи каталога по пути
  uint32_t GetDirEntryByPath(
    const wchar_t *pwszPath,
    uint32_t dwBaseStorageID,
    bool bUseAliases
    ) const;
  // Поиск записи каталога
  uint32_t FindDirEntry(
    uint32_t dwStorageID,
    uint8_t bObjectType,
    const wchar_t *pwszName,
    bool bUseAliases
    ) const;
  // Получение идентификатора родительской записи каталога
  uint32_t GetParentDirEntryID(
    uint32_t dwDirEntryID
    ) const;
  // Получение полного пути записи каталога
  size_t GetDirEntryFullPath(
    uint32_t dwDirEntryID,
    wchar_t *pBuffer,
    size_t nBufferSize,
    bool bUseAliases
    ) const;
  // Извлечение содержимого потока
  O2Res ExtractStreamData(
    uint32_t dwStreamID,
    uint64_t nOffset,
    size_t nSize,
    void *pBuffer,
    size_t *pnProcessedSize = NULL
    ) const;
  // Извлечение содержимого потока
  O2Res ExtractStreamData(
    uint32_t dwStreamID,
    uint64_t nOffset,
    uint64_t nSize,
    CSeqOutStream *pOutStream,
    uint64_t *pnProcessedSize = NULL
    ) const;
  // Извлечение потока
  O2Res ExtractStream(
    uint32_t dwStreamID,
    CSeqOutStream *pOutStream,
    uint64_t *pnProcessedSize = NULL
    ) const;

  // Конструктор класса
  COLE2File();
  // Деструктор класса
  virtual ~COLE2File();

private:
  // Копирование и присвоение не разрешены
  COLE2File(const COLE2File&);
  void operator=(const COLE2File&);

protected:
  // Проверка заголовка файла
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

  // Инициализация
  O2Res Init(
    CInStream *pInStream
    );
  // Поиск записи каталога
  uint32_t FindDirEntry(
    uint32_t dwStorageID,
    uint8_t bObjectType,
    const wchar_t *pwchName,
    size_t cchName,
    bool bUseAliases
    ) const;
  // Рекурсивная функция получения полного пути записи каталога
  size_t RecurseGetDirEntryFullPath(
    uint32_t dwDirEntryID,
    wchar_t *pBuffer,
    size_t nBufferSize,
    bool bUseAliases
    ) const;
  // Извлечение содержимого потока
  O2Res ExtractStreamData(
    uint32_t dwStreamStartSector,
    uint64_t nStreamSize,
    uint64_t nOffset,
    size_t nSize,
    void *pBuffer,
    size_t *pnProcessedSize
    ) const;
  // Извлечение содержимого потока
  O2Res ExtractStreamData(
    uint32_t dwStreamStartSector,
    uint64_t nStreamSize,
    uint64_t nOffset,
    uint64_t nSize,
    CSeqOutStream *pOutStream,
    uint64_t *pnProcessedSize
    ) const;
  // Выделение памяти и чтение FAT
  O2Res AllocAndReadFAT();
  // Выделение памяти и чтение mini FAT
  O2Res AllocAndReadMiniFAT();
  // Выделение памяти и чтение каталога
  O2Res AllocAndReadDirectory();
  // Создание списка элементов записей каталога
  O2Res CreateDirEntryItemList();
  // Рекурсивное перечисление записей каталога
  void EnumDirEntries(
    uint32_t dwDirEntryID,
    uint32_t dwParentDirEntryID
    );
  // Чтение данных mini-сектора в буфер
  O2Res ReadMiniSectorData(
    uint32_t dwMiniSectorNum,
    size_t nOffset,
    size_t nBytesToRead,
    void *pBuffer,
    size_t *pnBytesRead
    ) const;
  // Чтение сектора в буфер
  O2Res ReadSector(
    uint32_t dwSectorNum,
    void *pBuffer,
    size_t *pnBytesRead
    ) const;
  // Чтение данных сектора в буфер
  O2Res ReadSectorData(
    uint32_t dwSectorNum,
    size_t nOffset,
    size_t nBytesToRead,
    void *pBuffer,
    size_t *pnBytesRead
    ) const;
  // Получение номера следующего mini-сектора
  bool GetNextMiniSector(
    uint32_t dwMiniSectorNum,
    size_t nIndex,
    uint32_t *pdwMiniSectorNum
    ) const;
  // Получение номера следующего сектора
  bool GetNextSector(
    uint32_t dwSectorNum,
    size_t nIndex,
    uint32_t *pdwSectorNum
    ) const;
  // Проверка правильности номера сектора mini FAT
  bool IsMiniSectorNumValid(
    uint32_t dwMiniSectorNum
    ) const;
  // Проверка правильности номера сектора
  bool IsSectorNumValid(
    uint32_t dwSectorNum
    ) const;
  // Получение смещение сектора
  uint64_t GetSectorOffset(
    uint32_t dwSectorNum
    ) const;
};
//---------------------------------------------------------------------------
#endif  // __OLE2FILE_H__
