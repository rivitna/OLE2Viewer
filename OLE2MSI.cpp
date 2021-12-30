//---------------------------------------------------------------------------
#include "StrUtils.h"
#include "OLE2MSI.h"
//---------------------------------------------------------------------------
#ifndef countof
#define countof(a)  (sizeof(a)/sizeof(a[0]))
#endif  // countof
//---------------------------------------------------------------------------
#define MSI_NUM_BITS   6
#define MSI_NUM_CHARS  (1 << MSI_NUM_BITS)
#define MSI_CHAR_MASK  (MSI_NUM_CHARS - 1)
#define MSI_MIN_CHAR   0x3800
#define MSI_MAX_CHAR   (MSI_MIN_CHAR + (0x1000 | MSI_NUM_CHARS))

// Таблица символов имени записи каталога MSI
const char MSI_CODE_TABLE[MSI_NUM_CHARS] =
{
  '0', '1', '2', '3', '4', '5', '6', '7',
  '8', '9', 'A', 'B', 'C', 'D', 'E', 'F',
  'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N',
  'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V',
  'W', 'X', 'Y', 'Z', 'a', 'b', 'c', 'd',
  'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l',
  'm', 'n', 'o', 'p', 'q', 'r', 's', 't',
  'u', 'v', 'w', 'x', 'y', 'z', '.', '_'
};

// Наименования таблиц базы данных файла-контейнера MSI
const wchar_t* const MSI_TABLE_NAMES[] =
{
  L"_Columns",
  L"_Tables"
};

#define MSI_NUM_TABLE_NAMES  countof(MSI_TABLE_NAMES)
#define MSI_ALL_TABLES       ((1 << MSI_NUM_TABLE_NAMES) - 1)
//---------------------------------------------------------------------------
/***************************************************************************/
/* DecodeMsiDirEntryName - Декодирование имени записи каталога MSI         */
/***************************************************************************/
size_t DecodeMsiDirEntryName(
  const void *pDirEntryName,
  wchar_t *pwszDest,
  size_t cchDestLength
  )
{
  const uint16_t *pwchSrc = (const uint16_t *)pDirEntryName;
  wchar_t *pwchDest = pwszDest;

  if (!pwszDest)
    cchDestLength = 0;

  while (*pwchSrc)
  {
    if ((*pwchSrc < MSI_MIN_CHAR) || (*pwchSrc > MSI_MAX_CHAR))
      return 0;

    int c = *pwchSrc - MSI_MIN_CHAR;
    int c1 = c >> MSI_NUM_BITS;
    if (c1 <= MSI_NUM_CHARS)
    {
      if (cchDestLength != 0)
      {
        *pwchDest = (wchar_t)MSI_CODE_TABLE[c & MSI_CHAR_MASK];
        cchDestLength--;
      }
      pwchDest++;
      if (c1 == MSI_NUM_CHARS)
        break;
      if (cchDestLength != 0)
      {
        *pwchDest = (wchar_t)MSI_CODE_TABLE[c1];
        cchDestLength--;
      }
    }
    else
    {
      if (cchDestLength != 0)
      {
        *pwchDest = MSI_SPEC_CHAR;
        cchDestLength--;
      }
    }

    pwchDest++;
    pwchSrc++;
  }
  if (cchDestLength != 0)
    *pwchDest = L'\0';
  else
    pwchDest++;
  return (pwchDest - pwszDest);
}
/***************************************************************************/
/* DetectMsi - Детектирование файла-контейнера MSI                         */
/***************************************************************************/
bool DetectMsi(
  const COLE2File *pOLE2File
  )
{
  if (!pOLE2File || !pOLE2File->GetOpened())
    return false;

  // Поиск потоков с именами всех таблиц базы данных в корневом каталоге
  int nDetectedMSITables = 0;
  for (size_t i = 1; i < pOLE2File->GetDirEntryItemCount(); i++)
  {
    const DirEntryItem *pDirEntryItem;
    pDirEntryItem = pOLE2File->GetDirEntryItem(i);
    if ((pDirEntryItem->dwParentID != 0)  ||
        (pDirEntryItem->nType != DIR_ENTRY_TYPE_STREAM))
      continue;
    // Декодирование имени записи каталога MSI
    wchar_t wszMsiStreamName[MAX_MSI_DIR_ENTRY_NAME_LEN];
    size_t cchMsiStreamName;
    cchMsiStreamName = DecodeMsiDirEntryName(pDirEntryItem->pDirEntry->bName,
                                             wszMsiStreamName,
                                             countof(wszMsiStreamName));
    if ((cchMsiStreamName > 1) &&
        (cchMsiStreamName < countof(wszMsiStreamName)) &&
        (wszMsiStreamName[0] == MSI_SPEC_CHAR))
    {
      for (size_t j = 0; j < MSI_NUM_TABLE_NAMES; j++)
      {
        if ((nDetectedMSITables & (1 << j)) == 0)
        {
          // Сравнение имени потока с наименованием таблицы
          if (!StrCmpIW(&wszMsiStreamName[1], MSI_TABLE_NAMES[j]))
          {
            nDetectedMSITables |= (1 << j);
            if (nDetectedMSITables == MSI_ALL_TABLES)
              return true;
          }
        }
      }
    }
  }
  return false;
}
//---------------------------------------------------------------------------
