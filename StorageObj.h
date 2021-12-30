//---------------------------------------------------------------------------
#ifndef __STORAGEOBJ_H__
#define __STORAGEOBJ_H__
//---------------------------------------------------------------------------
#include <stddef.h>
#include "Streams.h"
#include "FileStreams.h"
#include "OLE2File.h"
//---------------------------------------------------------------------------
// Атрибуты для открываемых потоков
#define OPENABLE_STREAM_ATTRIBUTES  \
  (FILE_ATTRIBUTE_COMPRESSED | FILE_ATTRIBUTE_REPARSE_POINT)
//---------------------------------------------------------------------------
/***************************************************************************/
/*                                                                         */
/*                 CStorageObject - Класс файла-контейнера                 */
/*                                                                         */
/***************************************************************************/

// Тип файла OLE2
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

// Тип потока
typedef enum
{
  STREAM_UNKNOWN = 0,
  STREAM_VBA,
  STREAM_WORDDOC,
  STREAM_SUMINFO,
  STREAM_DOCSUMINFO,
  STREAM_NATIVEDATA
} StreamType;

// Информация о найденном элементе
typedef struct _FindItemData
{
  unsigned long     dwItemAttributes;
  FILETIME          ftCreationTime;
  FILETIME          ftModifiedTime;
  unsigned __int64  nItemSize;
  const wchar_t    *pwszItemName;
  unsigned long     dwItemID;
} FindItemData, *PFindItemData;

// Тип данных потока
typedef enum
{
  STREAM_DATA_TYPE_UNKNOWN = 0,
  STREAM_DATA_TYPE_VBA_MACRO_SRC,
  STREAM_DATA_TYPE_WORDDOC_TEXT,
  STREAM_DATA_TYPE_SUMINFO_TEXT,
  STREAM_DATA_TYPE_DOCSUMINFO_TEXT,
  STREAM_DATA_TYPE_NATIVE_DATA
} StreamDataType;

// Максимальная длина имени фиктивного элемента
#define MAX_DUMMY_ITEM_NAME_LEN  MAX_PATH

// Фиктивный элемент
typedef struct _DummyItem
{
  int      nItemType;
  size_t   nItemSize;
  wchar_t  wszItemName[MAX_DUMMY_ITEM_NAME_LEN];
} DummyItem, *PDummyItem;

class CStorageObject
{
public:
  // Флаг использования атрибутов для открываемых потоков
  bool UseOpenableStreamAttr;

  // Открытие
  int Open(
    const wchar_t *pwszFileName
    );
  // Закрытие
  void Close();
  // Получение флага открытия файла-контейнера
  inline bool GetOpened() const
    { return m_ole2File.GetOpened(); }
  // Получение пути к файлу-контейнеру
  inline const wchar_t *GetFilePath() const
    { return m_pwszFilePath; }
  // Получение имени файла-контейнера
  inline const wchar_t *GetFileName() const
    { return m_pwszFileName; }
  // Получение cтаршего номера версии
  inline int GetMajorVersion() const
    { return m_ole2File.GetMajorVersion(); }
  // Получение младшего номера версии
  inline int GetMinorVersion() const
    { return m_ole2File.GetMinorVersion(); }
  // Получение типа файла-контейнера
  inline int GetFileType() const
    { return m_nFileType; }
  // Получение количества потоков
  inline size_t GetStreamCount() const
    { return m_nNumStreams; }
  // Получение количества каталогов
  inline size_t GetStorageCount() const
    { return m_nNumStorages; }
  // Получение количества макросов
  inline size_t GetMacroCount() const
    { return m_nNumMacros; }
  // Получение размера файла
  unsigned __int64 GetFileSize() const;
  // Получение количества элементов в текущем каталоге
  size_t GetItemCount(
    bool bIncludeParentItem
    ) const;
  // Поиск первого элемента текущего каталога
  bool FindFirstItem(
    bool bIncludeParentItem,
    PFindItemData pFindItemData
    );
  // Поиск следующего элемента текущего каталога
  bool FindNextItem(
    PFindItemData pFindItemData
    );
  // Установка текущего каталога
  bool SetCurrentDir(
    const wchar_t *pwszDirPath
    );
  // Получение пути к текущему каталогу
  inline const wchar_t *GetCurrentDirPath() const
    { return m_pwszCurrentDirPath; }
  // Открытие потока
  bool OpenStream(
    unsigned long dwStreamID
    );
  // Извлечение элемента
  bool ExtractItem(
    unsigned long dwItemID,
    const wchar_t *pwszItemName,
    const wchar_t *pwszDestPath
    ) const;

  // Конструктор класса
  CStorageObject();
  // Деструктор класса
  ~CStorageObject();

private:
  // Копирование и присвоение не разрешены
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

  // Инициализация
  int Init(
    const wchar_t *pwszFileName
    );
  // Определение типа потока
  StreamType DetectStream(
    unsigned long dwStreamID
    ) const;
  // Извлечение элементов каталога
  bool ExtractStorageItems(
    unsigned long dwStorageID,
    const wchar_t *pwszDestPath
    ) const;
  // Извлечение потока
  bool ExtractStream(
    unsigned long dwStreamID,
    const wchar_t *pwszFileName
    ) const;
  // Извлечение данных потока
  bool ExtractStreamData(
    unsigned long dwStreamID,
    int nStreamDataType,
    const wchar_t *pwszFileName
    ) const;
  // Установка текущего родительского элемента каталога
  void SetCurParentItemID(
    unsigned long dwItemID,
    bool bIsStream
    );
  // Обновление пути к текущему каталогу
  void UpdateCurrentDirPath();
  // Определение типа файла-контейнере
  void DetectFileType();
  // Обновление информации о файле-контейнере
  void UpdateFileInfo();
  // Инициализация списка псевдонимов имен элементов
  void InitItemNameAliasList();
  // Поиск псевдонима имени элемента
  const wchar_t *FindItemNameAlias(
    unsigned long dwItemID
    ) const;
};
//---------------------------------------------------------------------------
// Преобразование типа файла OLE2 в строку
const wchar_t *OLE2FileTypeToString(
  int nFileType
  );
//---------------------------------------------------------------------------
#endif  // __STORAGEOBJ_H__
