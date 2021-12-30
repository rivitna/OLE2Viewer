//---------------------------------------------------------------------------
#include "OLE2Util.h"
//---------------------------------------------------------------------------
#define SPEC_MAX_CHAR  0x001F

// Таблица конвертации специальных символов ANSI/OEM в кодировку Unicode
const wchar_t SPEC_CHAR_CONVERT_TABLE[SPEC_MAX_CHAR + 1] =
{
  L'\x0000', L'\x263A', L'\x263B', L'\x2665',
  L'\x2666', L'\x2663', L'\x2660', L'\x2022',
  L'\x25D8', L'\x25CB', L'\x25D9', L'\x2642',
  L'\x2640', L'\x266A', L'\x266B', L'\x263C',
  L'\x25BA', L'\x25C4', L'\x2195', L'\x203C',
  L'\x00B6', L'\x00A7', L'\x25AC', L'\x21A8',
  L'\x2191', L'\x2193', L'\x2192', L'\x2190',
  L'\x221F', L'\x2194', L'\x25B2', L'\x25BC'
};
//---------------------------------------------------------------------------
/***************************************************************************/
/* IsDirEntryNameValid - Проверка, содержит ли имя записи каталога         */
/*                       недопустимые символы                              */
/***************************************************************************/
int IsDirEntryNameValid(
  const void *pDirEntryName
  )
{
  const uint16_t *pwch = (const uint16_t *)pDirEntryName;
  while (*pwch)
  {
    if (*pwch <= SPEC_MAX_CHAR)
      return 0;
    pwch++;
  }
  return 1;
}
/***************************************************************************/
/* DirEntryNameToFileName - Преобразование имени записи каталога в имя     */
/*                          файла                                          */
/***************************************************************************/
size_t DirEntryNameToFileName(
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
    if (cchDestLength != 0)
    {
      if (*pwchSrc <= SPEC_MAX_CHAR)
        *pwchDest = SPEC_CHAR_CONVERT_TABLE[*pwchSrc];
      else
        *pwchDest = (wchar_t)*pwchSrc;
      cchDestLength--;
    }
    pwchSrc++;
    pwchDest++;
  }
  if (cchDestLength != 0)
    *pwchDest = L'\0';
  else
    pwchDest++;
  return (pwchDest - pwszDest);
}
//---------------------------------------------------------------------------
