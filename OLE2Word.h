//---------------------------------------------------------------------------
#ifndef __OLE2WORD_H__
#define __OLE2WORD_H__
//---------------------------------------------------------------------------
#include "OLE2Type.h"
#include "OLE2Doc.h"
#include "OLE2File.h"
#include "Streams.h"
//---------------------------------------------------------------------------
/***************************************************************************/
/*                                                                         */
/*          COLE2WordDocumentStream - ����� ������ ��������� Word          */
/*                                    (������ "WordDocument")              */
/*                                                                         */
/***************************************************************************/

// Fib (��������� ���������)
#pragma pack(push, 1)
typedef struct _FIB
{
  FIBBASE    base;     // ������� ����� ��������� (FibBase)
  uint16_t   csw;      // ����� 16-������ �������� �� ������ ����� ���������
                       // fibRgW
                       // (=0x000E)
/*
  FIBRGW97   fibRgW;   // ������ ����� ��������� � 16-������� ����������
                       // (FibRgW97)
*/
  uint16_t   cslw;     // ����� 32-������ �������� � ������� ����� ���������
                       // fibRgLw
                       // (=0x0016)
  FIBRGLW97  fibRgLw;  // ������ ����� ��������� � 32-������� ����������
                       // (FibRgLW97)
} FIB, *PFIB;
#pragma pack(pop)

// ���������� � ����� ������ ���������
typedef struct _WDTextPiece
{
  size_t  fc;
  size_t  lcb;
  bool    bANSI;
} WDTextPiece, *PWDTextPiece;

class COLE2WordDocumentStream
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
  // ���������� ������
  O2Res ExtractText(
    CSeqOutStream *pOutStream,
    size_t *pnProcessedSize
    ) const;

  // ����������� ������
  COLE2WordDocumentStream();
  // ���������� ������
  virtual ~COLE2WordDocumentStream();

private:
  // ����������� � ���������� �� ���������
  COLE2WordDocumentStream(const COLE2WordDocumentStream&);
  void operator=(const COLE2WordDocumentStream&);

private:
  bool             m_bOpened;
  const COLE2File *m_pOLE2File;
  uint32_t         m_dwStreamID;
  size_t           m_nStreamSize;
  bool             m_bWord97;
  size_t           m_fcText;
  size_t           m_lcbText;
  PWDTextPiece     m_pTextPieces;
  size_t           m_nNumTextPieces;

  // �������������
  O2Res Init();
  // ����� ������ ������������ ���������
  O2Res ComplexRetrieveText(
    const FIB *pFib
    );
  // ����� ������ ������������ ���������
  O2Res ComplexRetrieveText(
    const FIB *pFib,
    const void *pClx,
    size_t cbClx
    );
  // ���������� ������ ������������ ���������
  O2Res ComplexExtractText(
    CSeqOutStream *pOutStream,
    size_t *pnProcessedSize
    ) const;
  // ���������� ������ (Word6)
  O2Res W6ExtractText(
    CSeqOutStream *pOutStream,
    size_t *pnProcessedSize
    ) const;
};
//---------------------------------------------------------------------------
// ���������� ������ �� ������ ��������� Word
O2Res ExtractWordDocumentText(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  );
//---------------------------------------------------------------------------
#endif  // __OLE2WORD_H__
