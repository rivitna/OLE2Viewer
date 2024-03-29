//---------------------------------------------------------------------------
#ifndef __OLE2NATIVE_H__
#define __OLE2NATIVE_H__
//---------------------------------------------------------------------------
#include "OLE2Type.h"
#include "OLE2File.h"
#include "Streams.h"
//---------------------------------------------------------------------------
/***************************************************************************/
/*                                                                         */
/*   COLE2NativeDataStream - ����� ������ ���������� ������� ����� OLE2    */
/*                           (������ "\1Ole10Native")                      */
/*                                                                         */
/***************************************************************************/

// ��� ������ ���������� �������
typedef enum
{
  NATIVE_DATA_UNKNOWN,
  NATIVE_DATA_PACKAGE,
  NATIVE_DATA_PBRUSH,
  NATIVE_DATA_WAVESOUND
} NativeDataType;

// ���������� � ��������� ������� (OLE Package)
typedef struct _OLEPackageInfo
{
  O2DataRange  LabelA;
  O2DataRange  LabelNameA;
  O2DataRange  FilePathA;
  O2DataRange  FileNameA;
  O2DataRange  TempPathA;
  O2DataRange  TempNameA;
  O2DataRange  LabelW;
  O2DataRange  LabelNameW;
  O2DataRange  FilePathW;
  O2DataRange  FileNameW;
  O2DataRange  TempPathW;
  O2DataRange  TempNameW;
} OLEPackageInfo, *POLEPackageInfo;

class COLE2NativeDataStream
{
public:
  // ��������
  O2Res Open(
    const COLE2File *pOLE2File,
    uint32_t dwStreamID
    );
  // ��������
  void Close();
  // ��������� ����� �������� ������
  inline bool GetOpened() const
    { return m_bOpened; }
  // ��������� ���� ������ ���������� �������
  inline NativeDataType GetNativeDataType() const
    { return m_nativeDataType; }
  // ��������� ������� ������ ���������� �������
  inline size_t GetNativeDataSize() const
    { return m_nativeData.nLength; }
  // ��������� ����� ���������� ������� (ANSI)
  inline size_t GetPackageLabelA(
    char *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringA(&m_pkgInfo.LabelA, pBuffer, nBufferSize); }
  // ��������� ����� ���������� ������� (Unicode)
  inline size_t GetPackageLabelW(
    wchar_t *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringW(&m_pkgInfo.LabelW, pBuffer, nBufferSize); }
  // ��������� ����� ���������� ������� (ANSI)
  inline size_t GetPackageLabelNameA(
    char *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringA(&m_pkgInfo.LabelNameA, pBuffer, nBufferSize); }
  // ��������� ����� ���������� ������� (Unicode)
  inline size_t GetPackageLabelNameW(
    wchar_t *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringW(&m_pkgInfo.LabelNameW, pBuffer, nBufferSize); }
  // ��������� ���� � ����� ���������� ������� (ANSI)
  inline size_t GetPackageFilePathA(
    char *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringA(&m_pkgInfo.FilePathA, pBuffer, nBufferSize); }
  // ��������� ���� � ����� ���������� ������� (Unicode)
  inline size_t GetPackageFilePathW(
    wchar_t *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringW(&m_pkgInfo.FilePathW, pBuffer, nBufferSize); }
  // ��������� ����� ����� ���������� ������� (ANSI)
  inline size_t GetPackageFileNameA(
    char *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringA(&m_pkgInfo.FileNameA, pBuffer, nBufferSize); }
  // ��������� ����� ����� ���������� ������� (Unicode)
  inline size_t GetPackageFileNameW(
    wchar_t *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringW(&m_pkgInfo.FileNameW, pBuffer, nBufferSize); }
  // ��������� ���� � ���������� ����� ���������� ������� (ANSI)
  inline size_t GetPackageTempPathA(
    char *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringA(&m_pkgInfo.TempPathA, pBuffer, nBufferSize); }
  // ��������� ���� � ���������� ����� ���������� ������� (Unicode)
  inline size_t GetPackageTempPathW(
    wchar_t *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringW(&m_pkgInfo.TempPathW, pBuffer, nBufferSize); }
  // ��������� ����� ���������� ����� ���������� ������� (ANSI)
  inline size_t GetPackageTempNameA(
    char *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringA(&m_pkgInfo.TempNameA, pBuffer, nBufferSize); }
  // ��������� ����� ���������� ����� ���������� ������� (Unicode)
  inline size_t GetPackageTempNameW(
    wchar_t *pBuffer,
    size_t nBufferSize
    ) const
    { return ExtractStringW(&m_pkgInfo.TempNameW, pBuffer, nBufferSize); }
  // ��������� ����� ����� ���������� �������
  size_t GetPackageFileName(
    wchar_t *pBuffer,
    size_t nBufferSize
    ) const;
  // ���������� ������ ���������� �������
  O2Res ExtractNativeData(
    CSeqOutStream *pOutStream,
    size_t *pnProcessedSize
    ) const;

  // ����������� ������
  COLE2NativeDataStream();
  // ���������� ������
  virtual ~COLE2NativeDataStream();

private:
  // ����������� � ���������� �� ���������
  COLE2NativeDataStream(const COLE2NativeDataStream&);
  void operator=(const COLE2NativeDataStream&);

private:
  bool             m_bOpened;
  NativeDataType   m_nativeDataType;
  O2DataRange      m_nativeData;
  OLEPackageInfo   m_pkgInfo;
  const COLE2File *m_pOLE2File;
  uint32_t         m_dwStreamID;
  size_t           m_nStreamDataSize;
  uint8_t         *m_pBuffer;
  size_t           m_nBufferSize;
  size_t           m_nBufferStartPos;
  size_t           m_nBytesRead;
  size_t           m_nCurrentPos;

  // �������� (OLE Package)
  O2Res OpenPackage();
  // ����������� ���� ������ ���������� �������
  O2Res DetectNativeDataType();
  // ���������� ������ (ANSI)
  size_t ExtractStringA(
    const O2DataRange *pStrRange,
    char *pBuffer,
    size_t nBufferSize
    ) const;
  // ���������� ������ (Unicode)
  size_t ExtractStringW(
    const O2DataRange *pStrRange,
    wchar_t *pBuffer,
    size_t nBufferSize
    ) const;
  // ������������ ���� � ����� (ANSI)
  O2Res ScanFilePathA(
    PO2DataRange pFilePathRange,
    PO2DataRange pFileNameRange
    );
  // ������������ ���� � ����� (Unicode)
  O2Res ScanFilePathW(
    PO2DataRange pFilePathRange,
    PO2DataRange pFileNameRange
    );
  // ������������ ������
  O2Res ScanData(
    void *buf,
    size_t count
    );
  // ������ ��������� ������ ������ � �����
  O2Res ReadNextStreamData();
  // ������ ������ ������ � �����
  O2Res ReadStreamData(
    size_t nStartPos
    );
};
//---------------------------------------------------------------------------
// ��������� ���������� � ��������� ������� �� ������ "\1Ole10Native"
O2Res GetEmbeddedObjectInfo(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  wchar_t *pwszFileNameBuf,
  size_t cchFileNameBuf,
  size_t *pnFileSize
  );
// ���������� ���������� ������� �� ������ "\1Ole10Native"
O2Res ExtractEmbeddedObject(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  );
//---------------------------------------------------------------------------
#endif  // __OLE2NATIVE_H__
