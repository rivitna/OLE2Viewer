//---------------------------------------------------------------------------
#define WIN32_LEAN_AND_MEAN  // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#include <malloc.h>
#include <memory.h>
#include <stdio.h>
#include <stdarg.h>
#include "OLE2CF.h"
#include "OLE2Metadata.h"
//---------------------------------------------------------------------------
// Идентификатор формата набора свойств Summary Information
// {F29F85E0-4FF9-1068-AB91-08002B27B3D9}
const uint8_t FMTID_SummaryInformation[16] =
{ 0xE0, 0x85, 0x9F, 0xF2, 0xF9, 0x4F, 0x68, 0x10, 0xAB, 0x91,
  0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9 };
// Идентификатор формата набора свойств Document Summary Information
// {D5CDD502-2E9C-101B-9397-08002B2CF9AE}
const uint8_t FMTID_DocSummaryInformation[16] =
{ 0x02, 0xD5, 0xCD, 0xD5, 0x9C, 0x2E, 0x1B, 0x10, 0x93, 0x97,
  0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE };
// Идентификатор формата определенного пользователем набора свойств потока
// DocumentSummaryInformation
// {D5CDD505-2E9C-101B-9397-08002B2CF9AE}
const uint8_t FMTID_UserDefinedProperties[16] =
{ 0x05, 0xD5, 0xCD, 0xD5, 0x9C, 0x2E, 0x1B, 0x10, 0x93, 0x97,
  0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE };

// Информация о свойствах набора FMTID_SummaryInformation
const PropertyInfo PSSI_PROP_INFOS[] =
{
  { PID_CODEPAGE,             PVT_I2,             "Code Page"           },
  { PID_LOCALEID,             PVT_UI4,            "Locale ID"           },
  { PIDSI_TITLE,              PVT_LPSTR,          "Title"               },
  { PIDSI_SUBJECT,            PVT_LPSTR,          "Subject"             },
  { PIDSI_AUTHOR,             PVT_LPSTR,          "Author"              },
  { PIDSI_KEYWORDS,           PVT_LPSTR,          "Keywords"            },
  { PIDSI_COMMENTS,           PVT_LPSTR,          "Comments"            },
  { PIDSI_TEMPLATE,           PVT_LPSTR,          "Template"            },
  { PIDSI_LASTAUTHOR,         PVT_LPSTR,          "Last Saved By"       },
  { PIDSI_REVNUMBER,          PVT_LPSTR,          "Revision Number"     },
  { PIDSI_EDITTIME,           PVT_FILETIME,       "Edit Time"           },
  { PIDSI_LASTPRINTED,        PVT_FILETIME,       "Last Printed"        },
  { PIDSI_CREATE_DTM,         PVT_FILETIME,       "Date Created"        },
  { PIDSI_LASTSAVE_DTM,       PVT_FILETIME,       "Date Last Saved"     },
  { PIDSI_PAGECOUNT,          PVT_I4,             "Page Count"          },
  { PIDSI_WORDCOUNT,          PVT_I4,             "Word Count"          },
  { PIDSI_CHARCOUNT,          PVT_I4,             "Character Count"     },
  { PIDSI_THUMBNAIL,          PVT_CF,             "Thumbnail"           },
  { PIDSI_APPNAME,            PVT_LPSTR,          "Application Name"    },
  { PIDSI_DOC_SECURITY,       PVT_I4,             "Security"            }
};

// Информация о свойствах набора FMTID_DocSummaryInformation
const PropertyInfo PSDSI_PROP_INFOS[] =
{
  { PID_CODEPAGE,             PVT_I2,             "Code Page"           },
  { PID_LOCALEID,             PVT_UI4,            "Locale ID"           },
  { PIDDSI_CATEGORY,          PVT_LPSTR,          "Category"            },
  { PIDDSI_PRESFORMAT,        PVT_LPSTR,          "Presentation Target" },
  { PIDDSI_BYTECOUNT,         PVT_I4,             "Byte Count"          },
  { PIDDSI_LINECOUNT,         PVT_I4,             "Line Count"          },
  { PIDDSI_PARACOUNT,         PVT_I4,             "Paragragh Count"     },
  { PIDDSI_SLIDECOUNT,        PVT_I4,             "Slides"              },
  { PIDDSI_NOTECOUNT,         PVT_I4,             "Notes"               },
  { PIDDSI_HIDDENCOUNT,       PVT_I4,             "Hidden Slides"       },
  { PIDDSI_MMCLIPCOUNT,       PVT_I4,             "Multimedia Clips"    },
  { PIDDSI_SCALE,             PVT_BOOL,           "Scale"               },
  { PIDDSI_HEADINGPAIR,       PVT_VARIANT_VECTOR, "Heading Pairs"       },
  { PIDDSI_DOCPARTS,          PVT_LPSTR_VECTOR,   "Document Parts"      },
  { PIDDSI_MANAGER,           PVT_LPSTR,          "Manager"             },
  { PIDDSI_COMPANY,           PVT_LPSTR,          "Company"             },
  { PIDDSI_LINKSDIRTY,        PVT_BOOL,           "Links Dirty"         },
  { PIDDSI_CCHWITHSPACES,     PVT_I4,             "Character Count"     },
  { PIDDSI_SHAREDDOC,         PVT_BOOL,           "Shared Document"     },
  { PIDDSI_HYPERLINKSCHANGED, PVT_BOOL,           "Hyperlinks Changed"  },
  { PIDDSI_VERSION,           PVT_I4,             "Application Version" },
  { PIDDSI_DIGSIG,            PVT_BLOB,           "Digital Signature"   },
  { PIDDSI_CONTENTTYPE,       PVT_LPSTR,          "Content Type"        },
  { PIDDSI_CONTENTSTATUS,     PVT_LPSTR,          "Content Status"      },
  { PIDDSI_LANGUAGE,          PVT_LPSTR,          "Language"            },
  { PIDDSI_DOCVERSION,        PVT_LPSTR,          "Document Version"    }
};
//---------------------------------------------------------------------------
/***************************************************************************/
/* CmpGUIDs - Сравнение GUID                                               */
/***************************************************************************/
#ifndef _WIN64
__declspec(naked)
#endif  // _WIN64
int __fastcall CmpGUIDs(
  const uint8_t guid1[16],
  const uint8_t guid2[16]
  )
{
#ifdef _WIN64
  return ((((uint64_t *)guid1)[0] == ((uint64_t *)guid2)[0]) &&
          (((uint64_t *)guid1)[1] == ((uint64_t *)guid2)[1]));
#else
  __asm
  {
    mov   eax,[ecx]
    cmp   eax,[edx]
    jne   NoEqual
    mov   eax,[ecx+4]
    cmp   eax,[edx+4]
    jne   NoEqual
    mov   eax,[ecx+8]
    cmp   eax,[edx+8]
    jne   NoEqual
    mov   eax,[ecx+12]
    sub   eax,[edx+12]
    jnz   NoEqual
    inc   eax
    ret

NoEqual:
    xor   eax,eax
    ret
  }
#endif  // _WIN64
}
/***************************************************************************/
/* GetPropVarSize - Получение размера вариантного значения свойства        */
/***************************************************************************/
size_t GetPropVarSize(
  const PROPERTY_VARIANT *pPropVar
  )
{
  switch (pPropVar->vt)
  {
    case PVT_I2:
    case PVT_I4:
    case PVT_BOOL:
    case PVT_UI4:
      return (2 * sizeof(uint32_t));
    case PVT_LPSTR:
    case PVT_BLOB:
    case PVT_CF:
      return ((sizeof(uint32_t) + sizeof(pPropVar->cb)) + pPropVar->cb);
    case PVT_FILETIME:
      return (sizeof(uint32_t) + sizeof(pPropVar->ftVal));
  }
  return 0;
}
/***************************************************************************/
/* FormatWriteStream - Форматированная запись в поток                      */
/***************************************************************************/
O2Res FormatWriteStream(
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize,
  const char *pszFormat,
  ...
  )
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  va_list arglist;
  int nLen;

  va_start(arglist, pszFormat);
  nLen = _vscprintf(pszFormat, arglist);
  va_end(arglist);
  if (nLen < 0)
    return O2_ERROR_DATA;
  if (nLen == 0)
    return O2_OK;

  if (!pOutStream)
  {
    if (pnProcessedSize)
      *pnProcessedSize = (size_t)nLen * sizeof(char);
    return O2_OK;
  }

  O2Res res = O2_OK;

  char buf[100];
  char *pBuf = buf;
  size_t nBufLen = sizeof(buf) / sizeof(char);

  if ((size_t)nLen >= nBufLen)
  {
    nBufLen = nLen + 1;
    pBuf = (char *)malloc(nBufLen * sizeof(char));
    if (!pBuf)
      return O2_ERROR_MEM;
  }

  va_start(arglist, pszFormat);
  nLen = vsprintf_s(pBuf, nBufLen, pszFormat, arglist);
  va_end(arglist);
  if (nLen > 0)
  {
    // Запись данных в поток
    size_t nBytesToWrite = (size_t)nLen * sizeof(char);
    size_t nBytesWritten = pOutStream->Write(pBuf, nBytesToWrite);
    if (nBytesWritten != nBytesToWrite)
      res = O2_ERROR_WRITE;
    if (pnProcessedSize)
      *pnProcessedSize = nBytesWritten;
  }
  else if (nLen < 0)
  {
    res = O2_ERROR_DATA;
  }

  if (pBuf != buf)
    free(pBuf);

  return res;
}
/***************************************************************************/
/* WriteStrVectorStream - Запись вектора строк в поток                     */
/***************************************************************************/
O2Res WriteStrVectorStream(
  const char *pszHeadingName,
  const void *pStrVector,
  size_t nItemCount,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  )
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  if (!pStrVector)
    nItemCount = 0;

  if (!pszHeadingName || !pszHeadingName[0])
  {
    if (nItemCount == 0)
      return O2_OK;
    pszHeadingName = "???";
  }

  O2Res res = O2_OK;

  size_t nProcessedSize;

  if (nItemCount == 0)
  {
    // Запись только названия в поток
    res = FormatWriteStream(pOutStream, &nProcessedSize, "%s:\r\n",
                            pszHeadingName);
    if (pnProcessedSize) *pnProcessedSize += nProcessedSize;
  }
  else if (nItemCount == 1)
  {
    // Запись в поток названия и одной строки
    res = FormatWriteStream(pOutStream, &nProcessedSize,
                            "%s: %s\r\n",
                            pszHeadingName,
                            (const char *)pStrVector + sizeof(uint32_t));
    if (pnProcessedSize) *pnProcessedSize += nProcessedSize;
  }
  else
  {
    // Запись названия в поток
    res = FormatWriteStream(pOutStream, &nProcessedSize, "%s (%u):\r\n",
                            pszHeadingName, (unsigned int)nItemCount);
    if (pnProcessedSize) *pnProcessedSize += nProcessedSize;

    // Запись строк вектора в поток
    const char *pStr = (const char *)pStrVector;
    do
    {
      uint32_t cch = ((uint32_t *)pStr)[0];
      pStr += sizeof(uint32_t);
      O2Res res2 = FormatWriteStream(pOutStream, &nProcessedSize, "%s\r\n",
                                     pStr);
      if ((res2 != O2_OK) && (res == O2_OK))
        res = res2;
      if (pnProcessedSize) *pnProcessedSize += nProcessedSize;
      pStr += cch;
      nItemCount--;
    }
    while (nItemCount != 0);
  }

  return res;
}
//---------------------------------------------------------------------------
/***************************************************************************/
/*                                                                         */
/*     COLE2PropertySetStream - класс потока набора свойств файла OLE2     */
/*                              (потоки "\5DocumentSummaryInformation",    */
/*                               "\5SummaryInformation")                   */
/*                                                                         */
/***************************************************************************/

/***************************************************************************/
/* COLE2PropertySetStream - Конструктор класса                             */
/***************************************************************************/
COLE2PropertySetStream::COLE2PropertySetStream()
  : m_bOpened(false),
    m_streamType(PSS_NONE),
    m_pOLE2File(NULL),
    m_dwStreamID(NOSTREAM),
    m_nStreamSize(0),
    m_nOSType(0),
    m_nOSMajorVersion(0),
    m_nOSMinorVersion(0),
    m_nNumPropSets(0),
    m_pStreamHeader(NULL),
    m_pPropInfos(NULL),
    m_nNumPropInfos(0)
{
  memset(&m_propSetBlobs, 0, sizeof(m_propSetBlobs));
}
/***************************************************************************/
/* ~COLE2PropertySetStream - Деструктор класса                             */
/***************************************************************************/
COLE2PropertySetStream::~COLE2PropertySetStream()
{
  // Закрытие
  Close();
}
/***************************************************************************/
/* Open - Открытие                                                         */
/***************************************************************************/
O2Res COLE2PropertySetStream::Open(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  PropertySetStreamType streamType
  )
{
  // Закрытие
  Close();

  if (!pOLE2File || !pOLE2File->GetOpened() ||
      (dwStreamID == NOSTREAM))
    return O2_ERROR_PARAM;

  const DIRECTORY_ENTRY *pDirEntry;
  pDirEntry = pOLE2File->GetDirEntry(dwStreamID);
  if (!pDirEntry ||
      (pDirEntry->bObjectType != OBJ_TYPE_STREAM))
    return O2_ERROR_PARAM;

  uint64_t nStreamSize = (pOLE2File->GetMajorVersion() > 3)
                           ? pDirEntry->qwStreamSize
                           : pDirEntry->dwLowStreamSize;
  if ((nStreamSize > V3_MAX_STREAM_SIZE) ||
      (pDirEntry->dwLowStreamSize < sizeof(PROPERTYSET_STREAM_HEADER)))
    return O2_ERROR_PARAM;

  m_pOLE2File = pOLE2File;
  m_dwStreamID = dwStreamID;
  m_streamType = streamType;
  m_nStreamSize = pDirEntry->dwLowStreamSize;

  // Инициализация
  O2Res res = Init();

  if (!m_bOpened)
  {
    // Закрытие
    Close();
  }

  return res;
}
/***************************************************************************/
/* Init - Инициализация                                                    */
/***************************************************************************/
O2Res COLE2PropertySetStream::Init()
{
  O2Res res;

  PROPERTYSET_STREAM_HEADER streamHeader;

  // Чтение заголовка потока без записей наборов свойств
  res = m_pOLE2File->ExtractStreamData(m_dwStreamID, 0,
                                       sizeof(PROPERTYSET_STREAM_HEADER),
                                       &streamHeader, NULL);
  if (res != O2_OK)
    return res;

  // Проверка заголовка потока
  if ((streamHeader.wVersion != 0) &&
      (streamHeader.wVersion != 1))
    return O2_ERROR_DATA;

  // Тип и версия ОС
  m_nOSType = streamHeader.wOSType;
  m_nOSMajorVersion = (uint8_t)streamHeader.wOSVersion;
  m_nOSMinorVersion = streamHeader.wOSVersion >> 8;

  // Количество прочитанных наборов свойств
  m_nNumPropSets = 0;

  // Количество наборов свойств
  size_t nNumPropSets = streamHeader.dwNumPropSets;
  if (nNumPropSets == 0)
    return O2_OK;

  // Определение размера заголовка потока
  size_t cbStreamHeader = sizeof(PROPERTYSET_STREAM_HEADER) +
                          nNumPropSets * sizeof(PROPERTYSET_ENTRY);
  if (cbStreamHeader > m_nStreamSize)
    return O2_ERROR_DATA;

  // Выделение памяти для заголовка потока
  m_pStreamHeader = (PPROPERTYSET_STREAM_HEADER)malloc(cbStreamHeader);
  if (!m_pStreamHeader)
    return O2_ERROR_MEM;

  // Чтение записей наборов свойств
  res = m_pOLE2File->ExtractStreamData(m_dwStreamID,
                                       sizeof(PROPERTYSET_STREAM_HEADER),
                                       nNumPropSets *
                                       sizeof(PROPERTYSET_ENTRY),
                                       (uint8_t *)m_pStreamHeader +
                                       sizeof(PROPERTYSET_STREAM_HEADER),
                                       NULL);
  if (res != O2_OK)
    return res;

  // Копирование прочитанной части заголовка потока
  memcpy(m_pStreamHeader, &streamHeader, sizeof(PROPERTYSET_STREAM_HEADER));

  switch (m_streamType)
  {
    case PSS_SUMMARYINFO:
      // Набор свойств Summary Information
      if (!CmpGUIDs(FMTID_SummaryInformation,
                    m_pStreamHeader->propSetEntries[0].bFMTID))
        return O2_ERROR_DATA;
      m_pPropInfos = PSSI_PROP_INFOS;
      m_nNumPropInfos = countof(PSSI_PROP_INFOS);
      break;
    case PSS_DOCSUMMARYINFO:
      // Набор свойств Document Summary Information
      if (!CmpGUIDs(FMTID_DocSummaryInformation,
                    m_pStreamHeader->propSetEntries[0].bFMTID))
        return O2_ERROR_DATA;
      m_pPropInfos = PSDSI_PROP_INFOS;
      m_nNumPropInfos = countof(PSDSI_PROP_INFOS);
      break;
    default:
      return O2_ERROR_DATA;
  }

  // Выделение памяти и чтение первого набора свойств
  res = AllocAndReadPropSet(m_pStreamHeader->propSetEntries[0].dwOffset,
                            &m_propSetBlobs[0]);
  if (!m_propSetBlobs[0].pPropSet)
  {
    if (res != O2_OK)
      return res;
  }

  m_nNumPropSets++;

  /* TODO: Набор свойств User Defined
  if ((nNumPropSets >= 2) &&
      (m_streamType == PSS_DOCSUMMARYINFO))
  {
    // Набор свойств User Defined
    if (CmpGUIDs(FMTID_UserDefinedProperties,
                 m_pStreamHeader->propSetEntries[1].bFMTID))
    {
      // Выделение памяти и чтение второго набора свойств
      O2Res res2;
      res2 = AllocAndReadPropSet(m_pStreamHeader->propSetEntries[1].dwOffset,
                                 &m_propSetBlobs[1]);
      if ((res2 != O2_OK) && (res == O2_OK))
        res = res2;
      if (m_propSetBlobs[1].pPropSet)
        m_nNumPropSets++;
    }
  }
  */

  m_bOpened = true;

  return res;
}
/***************************************************************************/
/* Close - Закрытие                                                        */
/***************************************************************************/
void COLE2PropertySetStream::Close()
{
  m_bOpened = false;
  m_streamType = PSS_NONE;
  m_pOLE2File = NULL;
  m_dwStreamID = NOSTREAM;
  m_nStreamSize = 0;
  m_nOSType = 0;
  m_nOSMajorVersion = 0;
  m_nOSMinorVersion = 0;
  m_nNumPropSets = 0;
  m_pPropInfos = NULL;
  m_nNumPropInfos = 0;
  // Освобождение выделенной памяти для данных наборов свойств
  for (size_t i = 0; i < countof(m_propSetBlobs); i++)
  {
    if (m_propSetBlobs[i].pPropSet)
    {
      free(m_propSetBlobs[i].pPropSet);
      m_propSetBlobs[i].pPropSet = NULL;
    }
    m_propSetBlobs[i].cbPropSet = 0;
  }
  if (m_pStreamHeader)
  {
    // Освобождение выделенной памяти для заголовка потока
    free(m_pStreamHeader);
    m_pStreamHeader = NULL;
  }
}
/***************************************************************************/
/* GetPropertySet - Получение набора свойств                               */
/***************************************************************************/
const PROPERTYSET_HEADER *COLE2PropertySetStream::GetPropertySet(
  size_t nIndex
  ) const
{
  if ((nIndex < m_nNumPropSets) &&
      m_propSetBlobs[nIndex].pPropSet &&
      (m_propSetBlobs[nIndex].cbPropSet != 0))
    return m_propSetBlobs[nIndex].pPropSet;
  return NULL;
}
/***************************************************************************/
/* GetPropertyData - Получение данных свойства                             */
/***************************************************************************/
O2Res COLE2PropertySetStream::GetPropertyData(
  size_t nIndex,
  PPropertyData pPropData
  ) const
{
  if (!pPropData)
    return O2_ERROR_PARAM;

  memset(pPropData, 0, sizeof(PropertyData));

  size_t cbPropSet = m_propSetBlobs[0].cbPropSet;
  PPROPERTYSET_HEADER pPropSet = m_propSetBlobs[0].pPropSet;

  if (!pPropSet || (cbPropSet == 0) ||
      (nIndex >= pPropSet->dwNumProperties))
    return O2_ERROR_PARAM;

  O2Res res = O2_OK;

  pPropData->dwID = pPropSet->propEntries[nIndex].dwID;

  // Поиск информации о свойстве по его идентификатору
  const PropertyInfo *pPropInfo;
  pPropInfo = FindPropInfoByPropID(pPropSet->propEntries[nIndex].dwID);
  if (pPropInfo)
    pPropData->pszName = pPropInfo->pszName;
  else
    res = O2_ERROR_DATA;

  // Получение смещения свойства и его проверка
  size_t nPropVarOffset = pPropSet->propEntries[nIndex].dwOffset;
  if ((nPropVarOffset < sizeof(PROPERTYSET_HEADER) +
                        pPropSet->dwNumProperties *
                        sizeof(PROPERTY_ENTRY)) ||
      (nPropVarOffset >= cbPropSet))
    return O2_ERROR_DATA;

  size_t nSize = cbPropSet - nPropVarOffset;
  if (nSize < MIN_PROP_VAR_SIZE)
    return O2_ERROR_DATA;

  PPROPERTY_VARIANT pPropVar;
  pPropVar = (PPROPERTY_VARIANT)((uint8_t *)pPropSet + nPropVarOffset);

  if (pPropInfo)
  {
    if (pPropVar->vt != pPropInfo->vt)
      res = O2_ERROR_DATA;
  }

  size_t nVarSize;
  size_t cch;

  if ((pPropVar->vt == PVT_VARIANT_VECTOR) ||
      (pPropVar->vt == PVT_LPSTR_VECTOR))
  {
    size_t nItemCount = pPropVar->nItemCount;
    nSize -= PROP_VAR_DATA_OFFSET;
    uint8_t *p = PROP_VAR_DATA(pPropVar);
    size_t i = 0;
    if (pPropVar->vt == PVT_VARIANT_VECTOR)
    {
      // Вектор вариантных значений
      for (; i < nItemCount; i++)
      {
        if (nSize < MIN_PROP_VAR_SIZE)
          break;
        PPROPERTY_VARIANT pVar = (PPROPERTY_VARIANT)p;
        // Получение размера вариантного значения свойства
        nVarSize = GetPropVarSize(pVar);
        if ((nVarSize == 0) || (nSize < nVarSize))
          break;
        nSize -= nVarSize;
        p += nVarSize;
        if (pVar->vt == PVT_LPSTR)
        {
          if (pVar->cb == 0)
            break;
          // Установка конечного нуль-символа строки
          if (p[-1] != '\0')
            p[-1] = '\0';
        }
      }
    }
    else
    {
      // Вектор строк
      for (; i < nItemCount; i++)
      {
        if (nSize < sizeof(uint32_t))
          break;
        nSize -= sizeof(uint32_t);
        cch = ((uint32_t *)p)[0];
        if ((cch == 0) || (nSize < cch))
          break;
        nSize -= cch;
        p += sizeof(uint32_t) + cch;
        // Установка конечного нуль-символа строки
        if (p[-1] != '\0')
          p[-1] = '\0';
      }
    }
    if (i < nItemCount)
    {
      nItemCount = i;
      res = O2_ERROR_DATA;
    }
    pPropData->nItemCount = nItemCount;
  }
  else if (pPropVar->vt == PVT_LPSTR)
  {
    // Проверка длины строки
    nSize -= PROP_VAR_DATA_OFFSET;
    cch = pPropVar->cb;
    if (nSize < cch)
    {
      cch = nSize;
      res = O2_ERROR_DATA;
    }
    if (cch == 0)
      return O2_ERROR_DATA;
    // Установка конечного нуль-символа строки
    char *pch = PROP_VAR_LPSTR(pPropVar);
    if (pch[cch - 1] != '\0')
      pch[cch - 1] = '\0';
  }
  else
  {
    // Получение размера вариантного значения свойства
    nVarSize = GetPropVarSize(pPropVar);
    if (nVarSize == 0)
      return O2_ERROR_DATA;
    if ((pPropVar->vt == PVT_I2) ||
        (pPropVar->vt == PVT_BOOL))
      nVarSize -= sizeof(uint16_t);
    if (nSize < nVarSize)
      return O2_ERROR_DATA;
  }

  pPropData->pVar = pPropVar;

  return res;
}
/***************************************************************************/
/* ExportToText - Экспорт в текстовой вид                                  */
/***************************************************************************/
O2Res COLE2PropertySetStream::ExportToText(
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  ) const
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  if (!m_bOpened)
    return O2_ERROR_PARAM;

  O2Res res = O2_OK;

  O2Res res2;
  size_t nProcessedSize;

  // Название набора свойств
  const char *pszPropertySetName;
  if (m_streamType == PSS_SUMMARYINFO)
    pszPropertySetName = "Summary Information";
  else if (m_streamType == PSS_DOCSUMMARYINFO)
    pszPropertySetName = "Document Summary Information";
  else
    pszPropertySetName = NULL;
  if (pszPropertySetName)
  {
    res2 = FormatWriteStream(pOutStream, &nProcessedSize,
                             "[%s]\r\n", pszPropertySetName);
    if ((res2 != O2_OK) && (res == O2_OK))
      res = res2;
    if (pnProcessedSize) *pnProcessedSize += nProcessedSize;
  }

  // Тип ОС
  const char* const OS_TYPES[] =
  {
    "Win16",
    "Mac",
    "Win32"
  };
  const char *pszOSType;
  pszOSType = ((m_nOSType >= 0) && (m_nOSType < countof(OS_TYPES)))
                ? OS_TYPES[m_nOSType] : "Unknown";
  res2 = FormatWriteStream(pOutStream, &nProcessedSize,
                           "OS Type: %d (%s)\r\n",
                           m_nOSType, pszOSType);
  if ((res2 != O2_OK) && (res == O2_OK))
    res = res2;
  if (pnProcessedSize) *pnProcessedSize += nProcessedSize;

  // Версия ОС
  res2 = FormatWriteStream(pOutStream, &nProcessedSize,
                           "OS Version: %d.%d\r\n",
                           m_nOSMajorVersion, m_nOSMinorVersion);
  if ((res2 != O2_OK) && (res == O2_OK))
    res = res2;
  if (pnProcessedSize) *pnProcessedSize += nProcessedSize;

  if (!m_propSetBlobs[0].pPropSet)
    return res;

  size_t nHeadingPairItemCount = 0;
  const PROPERTY_VARIANT *pHeadingPairVar = NULL;
  size_t nDocPartsItemCount = 0;
  const PROPERTY_VARIANT *pDocPartsVar = NULL;

  // Свойства
  for (size_t i = 0; i < m_propSetBlobs[0].pPropSet->dwNumProperties; i++)
  {
    PropertyData propData;

    // Получение данных свойства
    res2 = GetPropertyData(i, &propData);
    if (res2 != O2_OK)
    {
      if (res == O2_OK)
        res = res2;
      if (propData.pVar == NULL)
        continue;
    }

    if (m_streamType == PSS_DOCSUMMARYINFO)
    {
      switch (propData.dwID)
      {
        case PIDDSI_HEADINGPAIR:
          // Заголовки состава документа
          if (pHeadingPairVar)
          {
            if (res == O2_OK)
              res = O2_ERROR_DATA;
          }
          else if (propData.pVar->vt == PVT_VARIANT_VECTOR)
          {
            pHeadingPairVar = propData.pVar;
            nHeadingPairItemCount = propData.nItemCount;
            continue;
          }
          break;
        case PIDDSI_DOCPARTS:
          // Состав документа
          if (pDocPartsVar)
          {
            if (res == O2_OK)
              res = O2_ERROR_DATA;
          }
          else if (propData.pVar->vt == PVT_LPSTR_VECTOR)
          {
            pDocPartsVar = propData.pVar;
            nDocPartsItemCount = propData.nItemCount;
            continue;
          }
          break;
        case PIDDSI_DIGSIG:
          // Цифровая подпись проекта VBA
          if (propData.pVar->vt == PVT_BLOB)
            continue;
          break;
      }
    }
    else if (m_streamType == PSS_SUMMARYINFO)
    {
      if (propData.dwID == PIDSI_THUMBNAIL)
      {
        // Эскиз
        if (propData.pVar->vt == PVT_CF)
          continue;
      }
    }

    // Запись строки со свойством в поток
    res2 = WritePropLineStream(&propData, pOutStream, &nProcessedSize);
    if ((res2 != O2_OK) && (res == O2_OK))
      res = res2;
    if (pnProcessedSize) *pnProcessedSize += nProcessedSize;
  }

  if (pHeadingPairVar && (nHeadingPairItemCount != 0))
  {
    // Состав документа
    res2 = FormatWriteStream(pOutStream, &nProcessedSize,
                             "[%s]\r\n", "Document Parts");
    if ((res2 != O2_OK) && (res == O2_OK))
      res = res2;
    if (pnProcessedSize) *pnProcessedSize += nProcessedSize;

    const PROPERTY_VARIANT *pVar;
    const uint8_t *pStr;
    pVar = (PPROPERTY_VARIANT)PROP_VAR_DATA(pHeadingPairVar);
    pStr = (nDocPartsItemCount != 0) ? PROP_VAR_DATA(pDocPartsVar) : NULL;
    for (size_t i = 0; i < (nHeadingPairItemCount >> 1); i++)
    {
      const char *pszHeadingName = NULL;
      if (pVar->vt == PVT_LPSTR)
      {
        // Заголовок
        pszHeadingName = PROP_VAR_LPSTR(pVar);
      }
      pVar = (PPROPERTY_VARIANT)((uint8_t *)pVar + GetPropVarSize(pVar));
      if (pVar->vt == PVT_I4)
      {
        // Количество частей
        size_t nNumDocParts = pVar->dwVal;
        if (nNumDocParts > nDocPartsItemCount)
        {
          nNumDocParts = nDocPartsItemCount;
          if (res == O2_OK)
            res = O2_ERROR_DATA;
        }
        // Запись вектора строк в поток
        res2 = WriteStrVectorStream(pszHeadingName, pStr, nNumDocParts,
                                    pOutStream, &nProcessedSize);
        if ((res2 != O2_OK) && (res == O2_OK))
          res = res2;
        if (pnProcessedSize) *pnProcessedSize += nProcessedSize;
        nDocPartsItemCount -= nNumDocParts;
        while (nNumDocParts != 0)
        {
          pStr += sizeof(uint32_t) + ((uint32_t *)pStr)[0];
          nNumDocParts--;
        }
      }
      pVar = (PPROPERTY_VARIANT)((uint8_t *)pVar + GetPropVarSize(pVar));
    }
    if (nDocPartsItemCount != 0)
    {
      if (res == O2_OK)
        res = O2_ERROR_DATA;
    }
  }

  return res;
}
/***************************************************************************/
/* WritePropLineStream - Запись строки со свойством в поток                */
/***************************************************************************/
O2Res COLE2PropertySetStream::WritePropLineStream(
  const PropertyData *pPropData,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  ) const
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  O2Res res = O2_OK;

  const char *pszPropName;
  char szUnkPropName[32];

  // Получение имени свойства
  pszPropName = pPropData->pszName;
  if (!pszPropName)
  {
    res = O2_ERROR_DATA;
    // Получение имени для неизвестного свойства
    szUnkPropName[0] = '\0';
    sprintf_s(szUnkPropName, sizeof(szUnkPropName), "??? (PID=%08lX)",
              pPropData->dwID);
    pszPropName = szUnkPropName;
  }

  O2Res res2;

  switch (pPropData->pVar->vt)
  {
    case PVT_I2:
    case PVT_I4:
    case PVT_UI4:
      // Целочисленный тип
      {
        uint32_t dwVal;
        dwVal = (pPropData->pVar->vt == PVT_I2) ? pPropData->pVar->wVal
                                                : pPropData->pVar->dwVal;
        if (pPropData->dwID == PID_LOCALEID)
        {
          // Locale ID
          res2 = FormatWriteStream(pOutStream, pnProcessedSize,
                                   "%s: %08lX\r\n",
                                   pszPropName, dwVal);
        }
        else if ((m_streamType == PSS_DOCSUMMARYINFO) &&
                 (pPropData->dwID == PIDDSI_VERSION))
        {
          // Версия приложения
          res2 = FormatWriteStream(pOutStream, pnProcessedSize,
                                   "%s: %d.%d\r\n",
                                   pszPropName,
                                   (int)(uint16_t)(dwVal >> 16),
                                   (int)(uint16_t)dwVal);
        }
        else if ((m_streamType == PSS_SUMMARYINFO) &&
                 (pPropData->dwID == PIDSI_DOC_SECURITY))
        {
          // Набор флагов контроля доступа
          res2 = FormatWriteStream(pOutStream, pnProcessedSize,
                                   "%s: %08lX\r\n",
                                   pszPropName, dwVal);
        }
        else
        {
          res2 = FormatWriteStream(pOutStream, pnProcessedSize,
                                   "%s: %lu\r\n",
                                   pszPropName, dwVal);
        }
      }
      break;

    case PVT_BOOL:
      // Логический тип
      res2 = FormatWriteStream(pOutStream, pnProcessedSize,
                               "%s: %s\r\n",
                               pszPropName,
                               (pPropData->pVar->bVal == 0) ? "False"
                                                            : "True");
      break;

    case PVT_LPSTR:
      // Строковый тип
      res2 = FormatWriteStream(pOutStream, pnProcessedSize,
                               "%s: %s\r\n",
                               pszPropName,
                               PROP_VAR_LPSTR(pPropData->pVar));
      break;

    case PVT_FILETIME:
      // Тип времени файла
      if ((m_streamType == PSS_SUMMARYINFO) &&
          (pPropData->dwID == PIDSI_EDITTIME))
      {
        // Время редактирования
        unsigned int msecs;
        msecs = (pPropData->pVar->qwVal % 10000000 + 5000) / 10000;
        uint64_t secs = pPropData->pVar->qwVal / 10000000;
        unsigned int seconds = (unsigned int)(secs % 60);
        unsigned int minutes = (unsigned int)((secs / 60) % 60);
        unsigned int hours = (unsigned int)((secs / 3600) % 24);
        unsigned int days = (unsigned int)(secs / (24 * 3600));
        if (days == 0)
        {
          res2 = FormatWriteStream(pOutStream, pnProcessedSize,
                                   "%s: %02u:%02u:%02u.%03u\r\n",
                                   pszPropName,
                                   hours, minutes, seconds, msecs);
        }
        else
        {
          res2 = FormatWriteStream(pOutStream, pnProcessedSize,
                                   "%s: %ud %02u:%02u:%02u.%03u\r\n",
                                   pszPropName,
                                   days, hours, minutes, seconds, msecs);
        }
      }
      else
      {
        if ((pPropData->pVar->ftVal.dwLowDateTime != 0) ||
            (pPropData->pVar->ftVal.dwHighDateTime != 0))
        {
          SYSTEMTIME st;
          if (::FileTimeToSystemTime((FILETIME *)&pPropData->pVar->ftVal,
                                     &st))
          {
            res2 = FormatWriteStream(pOutStream, pnProcessedSize,
                                     "%s: %04u/%02u/%02u " \
                                     "%02u:%02u:%02u.%03u (UTC)\r\n",
                                     pszPropName,
                                     (unsigned int)st.wYear,
                                     (unsigned int)st.wMonth,
                                     (unsigned int)st.wDay,
                                     (unsigned int)st.wHour,
                                     (unsigned int)st.wMinute,
                                     (unsigned int)st.wSecond,
                                     (unsigned int)st.wMilliseconds);
          }
        }
      }
      break;

    default:
      // Свойство со значением неизвестного типа
      res = O2_ERROR_DATA;
      res2 = FormatWriteStream(pOutStream, pnProcessedSize,
                               "%s: ??? (Type=%04X)\r\n",
                               pszPropName,
                               (unsigned int)pPropData->pVar->vt);
      break;
  }

  if ((res2 != O2_OK) && (res == O2_OK))
    res = res2;
  return res;
}
/***************************************************************************/
/* AllocAndReadPropSet - Выделение памяти и чтение набора свойств          */
/***************************************************************************/
O2Res COLE2PropertySetStream::AllocAndReadPropSet(
  size_t nPropSetOffset,
  PPropertySetBlob pPropSetBlob
  ) const
{
  if ((nPropSetOffset < sizeof(PROPERTYSET_STREAM_HEADER) +
                        m_pStreamHeader->dwNumPropSets *
                        sizeof(PROPERTYSET_ENTRY)) ||
      (nPropSetOffset >= m_nStreamSize) ||
      (m_nStreamSize - nPropSetOffset < sizeof(PROPERTYSET_HEADER)))
    return O2_ERROR_DATA;

  O2Res res;

  PROPERTYSET_HEADER psHeader;

  // Чтение заголовка набора свойств без записей свойств
  res = m_pOLE2File->ExtractStreamData(m_dwStreamID,
                                       nPropSetOffset,
                                       sizeof(PROPERTYSET_HEADER),
                                       &psHeader, NULL);
  if (res != O2_OK)
    return res;

  if (psHeader.dwNumProperties == 0)
    return O2_OK;

  PPROPERTYSET_HEADER pPropSet;
  size_t cbPropSet;

  cbPropSet = psHeader.dwSize;
  if (cbPropSet > m_nStreamSize - nPropSetOffset)
  {
    res = O2_ERROR_DATA;
    cbPropSet = m_nStreamSize - nPropSetOffset;
  }
  if (cbPropSet < sizeof(PROPERTYSET_HEADER) +
                  psHeader.dwNumProperties * sizeof(PROPERTY_ENTRY))
    return O2_ERROR_DATA;

  // Выделение памяти для набора свойств
  pPropSet = (PPROPERTYSET_HEADER)malloc(cbPropSet);
  if (!pPropSet)
    return O2_ERROR_MEM;

  size_t nBytesRead;

  // Чтение оставшейся части набора свойств
  O2Res res2 = m_pOLE2File->ExtractStreamData(m_dwStreamID,
                                              nPropSetOffset +
                                              sizeof(PROPERTYSET_HEADER),
                                              cbPropSet -
                                              sizeof(PROPERTYSET_HEADER),
                                              (uint8_t *)pPropSet +
                                              sizeof(PROPERTYSET_HEADER),
                                              &nBytesRead);
  if ((res2 != O2_OK) && (res == O2_OK))
    res = res2;
  if (nBytesRead < psHeader.dwNumProperties * sizeof(PROPERTY_ENTRY))
  {
    free(pPropSet);
    return (res != O2_OK) ? res : O2_ERROR_DATA;
  }

  // Копирование прочитанной части заголовка набора свойств
  pPropSet->dwSize = psHeader.dwSize;
  pPropSet->dwNumProperties = psHeader.dwNumProperties;

  pPropSetBlob->pPropSet = pPropSet;
  pPropSetBlob->cbPropSet = sizeof(PROPERTYSET_HEADER) + nBytesRead;
  if (cbPropSet != pPropSetBlob->cbPropSet)
  {
    if (res == O2_OK)
      res = O2_ERROR_DATA;
  }

  return res;
}
/***************************************************************************/
/* FindPropInfoByPropID - Поиск информации о свойстве по его идентификатору*/
/***************************************************************************/
const PropertyInfo *COLE2PropertySetStream::FindPropInfoByPropID(
  uint32_t dwPropID
  ) const
{
  const PropertyInfo *pPropInfo = m_pPropInfos;
  for (size_t i = 0; i < m_nNumPropInfos; i++, pPropInfo++)
  {
    if (dwPropID == pPropInfo->dwID)
      return pPropInfo;
  }
  return NULL;
}
//---------------------------------------------------------------------------
/***************************************************************************/
/* ExtractMetadata - Извлечение метаданных из потоков                      */
/*                   "\5SummaryInformation", "\5DocumentSummaryInformation"*/
/***************************************************************************/
O2Res ExtractMetadata(
  const COLE2File *pOLE2File,
  uint32_t dwStreamID,
  PropertySetStreamType streamType,
  CSeqOutStream *pOutStream,
  size_t *pnProcessedSize
  )
{
  if (pnProcessedSize)
    *pnProcessedSize = 0;

  O2Res res;

  // Создание объекта потока набора свойств файла OLE2
  COLE2PropertySetStream psStream;

  // Открытие потока набора свойств
  res = psStream.Open(pOLE2File, dwStreamID, streamType);
  if (psStream.GetOpened())
  {
    // Экспорт в текстовой вид
    O2Res res2 = psStream.ExportToText(pOutStream, pnProcessedSize);
    if ((res2 != O2_OK) && (res == O2_OK))
      res = res2;
  }

  return res;
}
//---------------------------------------------------------------------------
