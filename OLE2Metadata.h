//---------------------------------------------------------------------------
#ifndef __OLE2METADATA_H__
#define __OLE2METADATA_H__
//---------------------------------------------------------------------------
#include "OLE2Type.h"
#include "OLE2PS.h"
#include "OLE2File.h"
#include "Streams.h"
//---------------------------------------------------------------------------
/***************************************************************************/
/*                                                                         */
/*     COLE2PropertySetStream - класс потока набора свойств файла OLE2     */
/*                              (потоки "\5DocumentSummaryInformation",    */
/*                               "\5SummaryInformation")                   */
/*                                                                         */
/***************************************************************************/

// Тип потока набора свойств
typedef enum
{
  PSS_NONE,
  PSS_SUMMARYINFO,
  PSS_DOCSUMMARYINFO
} PropertySetStreamType;

// Данные свойства
typedef struct _PropertyData
{
  uint32_t                dwID;
  const char             *pszName;
  size_t                  nItemCount;
  const PROPERTY_VARIANT *pVar;
} PropertyData, *PPropertyData;

// Информация о свойстве
typedef struct _PropertyInfo
{
  uint32_t    dwID;
  uint16_t    vt;
  const char *pszName;
} PropertyInfo, *PPropertyInfo;

// Данные набора свойств
typedef struct _PropertySetBlob
{
  size_t               cbPropSet;
  PPROPERTYSET_HEADER  pPropSet;
} PropertySetBlob, *PPropertySetBlob;

class COLE2PropertySetStream
{
public:
  // Открытие
  O2Res Open(
    const COLE2File *pOLE2File,
    uint32_t dwStreamID,
    PropertySetStreamType streamType
    );
  // Закрытие
  void Close();
  // Получение флага открытия потока
  inline bool GetOpened() const
    { return m_bOpened; }
  // Получение типа ОС
  inline int GetOSType() const
    { return m_nOSType; }
  // Получение cтаршего номера версии ОС
  inline int GetOSMajorVersion() const
    { return m_nOSMajorVersion; }
  // Получение младшего номера версии ОС
  inline int GetOSMinorVersion() const
    { return m_nOSMinorVersion; }
  // Получение заголовка потока
  inline const PROPERTYSET_STREAM_HEADER *GetStreamHeader() const
    { return m_pStreamHeader; }
  // Получение количества наборов свойств
  inline size_t GetPropSetCount() const
    { return m_nNumPropSets; }
  // Получение набора свойств
  const PROPERTYSET_HEADER *GetPropertySet(
    size_t nIndex
    ) const;
  // Получение данных свойства
  O2Res GetPropertyData(
    size_t nIndex,
    PPropertyData pPropData
    ) const;
  // Экспорт в текстовой вид
  O2Res ExportToText(
    CSeqOutStream *pOutStream,
    size_t *pnProcessedSize = NULL
    ) const;

  // Конструктор класса
  COLE2PropertySetStream();
  // Деструктор класса
  virtual ~COLE2PropertySetStream();

private:
  // Копирование и присвоение не разрешены
  COLE2PropertySetStream(const COLE2PropertySetStream&);
  void operator=(const COLE2PropertySetStream&);

private:
  bool                        m_bOpened;
  PropertySetStreamType       m_streamType;
  const COLE2File            *m_pOLE2File;
  uint32_t                    m_dwStreamID;
  size_t                      m_nStreamSize;
  int                         m_nOSType;
  int                         m_nOSMajorVersion;
  int                         m_nOSMinorVersion;
  size_t                      m_nNumPropSets;
  PPROPERTYSET_STREAM_HEADER  m_pStreamHeader;
  PropertySetBlob             m_propSetBlobs[2];
  const PropertyInfo         *m_pPropInfos;
  size_t                      m_nNumPropInfos;

  // Инициализация
  O2Res Init();
  // Запись строки со свойством в поток
  O2Res WritePropLineStream(
    const PropertyData *pPropData,
    CSeqOutStream *pOutStream,
    size_t *pnProcessedSize
    ) const;
  // Выделение памяти и чтение набора свойств
  O2Res AllocAndReadPropSet(
    size_t nPropSetOffset,
    PPropertySetBlob pPropSetBlob
    ) const;
  // Поиск информации о свойстве по его идентификатору
  const PropertyInfo *FindPropInfoByPropID(
    uint32_t dwPropID
    ) const;
};
//---------------------------------------------------------------------------
// Извлечение метаданных из потоков "\5SummaryInformation",
// "\5DocumentSummaryInformation"
O2Res ExtractMetadata(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  PropertySetStreamType streamType,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  );
//---------------------------------------------------------------------------
#endif  // __OLE2METADATA_H__
