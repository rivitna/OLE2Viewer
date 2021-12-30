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
/*          COLE2WordDocumentStream - класс потока документа Word          */
/*                                    (потока "WordDocument")              */
/*                                                                         */
/***************************************************************************/

// Fib (Заголовок документа)
#pragma pack(push, 1)
typedef struct _FIB
{
  FIBBASE    base;     // Базовая часть заголовка (FibBase)
  uint16_t   csw;      // Число 16-битных значений во второй части заголовка
                       // fibRgW
                       // (=0x000E)
/*
  FIBRGW97   fibRgW;   // Вторая часть заголовка с 16-битными значениями
                       // (FibRgW97)
*/
  uint16_t   cslw;     // Число 32-битных значений в третьей части заголовка
                       // fibRgLw
                       // (=0x0016)
  FIBRGLW97  fibRgLw;  // Третья часть заголовка с 32-битными значениями
                       // (FibRgLW97)
} FIB, *PFIB;
#pragma pack(pop)

// Информация о части текста документа
typedef struct _WDTextPiece
{
  size_t  fc;
  size_t  lcb;
  bool    bANSI;
} WDTextPiece, *PWDTextPiece;

class COLE2WordDocumentStream
{
public:
  // Открытие
  O2Res Open(
    const COLE2File *pOLE2File,
    uint32_t dwStreamID
    );
  // Закрытие
  void Close();
  // Получение флага открытия потока
  inline bool GetOpened() const
    { return m_bOpened; }
  // Извлечение текста
  O2Res ExtractText(
    CSeqOutStream *pOutStream,
    size_t *pnProcessedSize
    ) const;

  // Конструктор класса
  COLE2WordDocumentStream();
  // Деструктор класса
  virtual ~COLE2WordDocumentStream();

private:
  // Копирование и присвоение не разрешены
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

  // Инициализация
  O2Res Init();
  // Поиск текста комплексного документа
  O2Res ComplexRetrieveText(
    const FIB *pFib
    );
  // Поиск текста комплексного документа
  O2Res ComplexRetrieveText(
    const FIB *pFib,
    const void *pClx,
    size_t cbClx
    );
  // Извлечение текста комплексного документа
  O2Res ComplexExtractText(
    CSeqOutStream *pOutStream,
    size_t *pnProcessedSize
    ) const;
  // Извлечение текста (Word6)
  O2Res W6ExtractText(
    CSeqOutStream *pOutStream,
    size_t *pnProcessedSize
    ) const;
};
//---------------------------------------------------------------------------
// Извлечение текста из потока документа Word
O2Res ExtractWordDocumentText(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  );
//---------------------------------------------------------------------------
#endif  // __OLE2WORD_H__
