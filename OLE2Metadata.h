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
/*     COLE2PropertySetStream - ����� ������ ������ ������� ����� OLE2     */
/*                              (������ "\5DocumentSummaryInformation",    */
/*                               "\5SummaryInformation")                   */
/*                                                                         */
/***************************************************************************/

// ��� ������ ������ �������
typedef enum
{
  PSS_NONE,
  PSS_SUMMARYINFO,
  PSS_DOCSUMMARYINFO
} PropertySetStreamType;

// ������ ��������
typedef struct _PropertyData
{
  uint32_t                dwID;
  const char             *pszName;
  size_t                  nItemCount;
  const PROPERTY_VARIANT *pVar;
} PropertyData, *PPropertyData;

// ���������� � ��������
typedef struct _PropertyInfo
{
  uint32_t    dwID;
  uint16_t    vt;
  const char *pszName;
} PropertyInfo, *PPropertyInfo;

// ������ ������ �������
typedef struct _PropertySetBlob
{
  size_t               cbPropSet;
  PPROPERTYSET_HEADER  pPropSet;
} PropertySetBlob, *PPropertySetBlob;

class COLE2PropertySetStream
{
public:
  // ��������
  O2Res Open(
    const COLE2File *pOLE2File,
    uint32_t dwStreamID,
    PropertySetStreamType streamType
    );
  // ��������
  void Close();
  // ��������� ����� �������� ������
  inline bool GetOpened() const
    { return m_bOpened; }
  // ��������� ���� ��
  inline int GetOSType() const
    { return m_nOSType; }
  // ��������� c������� ������ ������ ��
  inline int GetOSMajorVersion() const
    { return m_nOSMajorVersion; }
  // ��������� �������� ������ ������ ��
  inline int GetOSMinorVersion() const
    { return m_nOSMinorVersion; }
  // ��������� ��������� ������
  inline const PROPERTYSET_STREAM_HEADER *GetStreamHeader() const
    { return m_pStreamHeader; }
  // ��������� ���������� ������� �������
  inline size_t GetPropSetCount() const
    { return m_nNumPropSets; }
  // ��������� ������ �������
  const PROPERTYSET_HEADER *GetPropertySet(
    size_t nIndex
    ) const;
  // ��������� ������ ��������
  O2Res GetPropertyData(
    size_t nIndex,
    PPropertyData pPropData
    ) const;
  // ������� � ��������� ���
  O2Res ExportToText(
    CSeqOutStream *pOutStream,
    size_t *pnProcessedSize = NULL
    ) const;

  // ����������� ������
  COLE2PropertySetStream();
  // ���������� ������
  virtual ~COLE2PropertySetStream();

private:
  // ����������� � ���������� �� ���������
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

  // �������������
  O2Res Init();
  // ������ ������ �� ��������� � �����
  O2Res WritePropLineStream(
    const PropertyData *pPropData,
    CSeqOutStream *pOutStream,
    size_t *pnProcessedSize
    ) const;
  // ��������� ������ � ������ ������ �������
  O2Res AllocAndReadPropSet(
    size_t nPropSetOffset,
    PPropertySetBlob pPropSetBlob
    ) const;
  // ����� ���������� � �������� �� ��� ��������������
  const PropertyInfo *FindPropInfoByPropID(
    uint32_t dwPropID
    ) const;
};
//---------------------------------------------------------------------------
// ���������� ���������� �� ������� "\5SummaryInformation",
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
