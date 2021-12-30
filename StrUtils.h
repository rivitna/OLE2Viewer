//---------------------------------------------------------------------------
#ifndef __STRUTILS_H__
#define __STRUTILS_H__
//---------------------------------------------------------------------------
#include <tchar.h>
//---------------------------------------------------------------------------
#ifdef __cplusplus
extern "C" {
#endif  // __cplusplus
//---------------------------------------------------------------------------
#ifdef _UNICODE
#define SkipBlanks     SkipBlanksA
#define StrChar        StrCharW
#define StrCharSet     StrCharSetW
#define StrRChar       StrRCharW
#define StrRCharSet    StrRCharSetW
#define StrCmpNI       StrCmpNIW
#define StrCmpI        StrCmpIW
#define StrToULN       StrToULNW
#define StrToUL        StrToULW
#define StrToUI        StrToUIW
#define UI64ToStr      UI64ToStrW
#define BinToHexStr    BinToHexStrW
#define HexDigitValue  HexDigitValueW
#define HexStrToBin    HexStrToBinW
#else
#define SkipBlanks     SkipBlanksW
#define StrChar        StrCharA
#define StrCharSet     StrCharSetA
#define StrRChar       StrRCharA
#define StrRCharSet    StrRCharSetA
#define StrCmpNI       StrCmpNIA
#define StrCmpI        StrCmpIA
#define StrToULN       StrToULNA
#define StrToUL        StrToULA
#define StrToUI        StrToUIA
#define UI64ToStr      UI64ToStrA
#define BinToHexStr    BinToHexStrA
#define HexDigitValue  HexDigitValueA
#define HexStrToBin    HexStrToBinA
#endif  // _UNICODE
//---------------------------------------------------------------------------
// Пропуск начальных пробелов в строке
char *SkipBlanksA(
  const char *s
  );
// Пропуск начальных пробелов в строке
wchar_t *SkipBlanksW(
  const wchar_t *s
  );
// Поиск в строке первого символа
char *StrCharA(
  const char *s,
  char ch
  );
// Поиск в строке первого символа
wchar_t *StrCharW(
  const wchar_t *s,
  wchar_t ch
  );
// Поиск в строке первого символа из набора
char *StrCharSetA(
  const char *s,
  const char *cs
  );
// Поиск в строке первого символа из набора
wchar_t *StrCharSetW(
  const wchar_t *s,
  const wchar_t *cs
  );
// Поиск в строке последнего символа
char *StrRCharA(
  const char *s,
  char ch
  );
// Поиск в строке последнего символа
wchar_t *StrRCharW(
  const wchar_t *s,
  wchar_t ch
  );
// Поиск в строке последнего символа из набора
char *StrRCharSetA(
  const char *s,
  const char *cs
  );
// Поиск в строке последнего символа из набора
wchar_t *StrRCharSetW(
  const wchar_t *s,
  const wchar_t *cs
  );
// Сравнение строк без учета регистра
int StrCmpNIA(
  const char *s1,
  const char *s2,
  size_t count
  );
// Сравнение строк без учета регистра
int StrCmpNIW(
  const wchar_t *s1,
  const wchar_t *s2,
  size_t count
  );
// Сравнение строк без учета регистра
int StrCmpIA(
  const char *s1,
  const char *s2
  );
// Сравнение строк без учета регистра
int StrCmpIW(
  const wchar_t *s1,
  const wchar_t *s2
  );
// Преобразование строки в беззнаковое целое
unsigned long StrToULNA(
  const char *s,
  size_t count,
  char **endptr,
  int base
  );
// Преобразование строки в беззнаковое целое
unsigned long StrToULNW(
  const wchar_t *s,
  size_t count,
  wchar_t **endptr,
  int base
  );
// Преобразование строки в беззнаковое целое
unsigned long StrToULA(
  const char *s,
  char **endptr,
  int base
  );
// Преобразование строки в беззнаковое целое
unsigned long StrToULW(
  const wchar_t *s,
  wchar_t **endptr,
  int base
  );
// Преобразование строки в беззнаковое целое
unsigned int StrToUIA(
  const char *s
  );
// Преобразование строки в беззнаковое целое
unsigned int StrToUIW(
  const wchar_t *s
  );
// Преобразование беззнакового целого в строку
size_t UI64ToStrA(
  unsigned __int64 nValue,
  char chThousandSeparator,
  char *pBuffer
  );
// Преобразование беззнакового целого в строку
size_t UI64ToStrW(
  unsigned __int64 nValue,
  wchar_t chThousandSeparator,
  wchar_t *pBuffer
  );
// Преобразование бинарных данных в шестнадцатиричную строку
void BinToHexStrA(
  const void *pData,
  size_t cbData,
  char *pBuffer
  );
// Преобразование бинарных данных в шестнадцатиричную строку
void BinToHexStrW(
  const void *pData,
  size_t cbData,
  wchar_t *pBuffer
  );
// Преобразование шестнадцатиричной цифры в число
int HexDigitValueA(
  char ch
  );
// Преобразование шестнадцатиричной цифры в число
int HexDigitValueW(
  wchar_t ch
  );
// Преобразование шестнадцатиричной строки в бинарные данные
size_t HexStrToBinA(
  const char *s,
  void *pBuffer,
  size_t cbSize
  );
// Преобразование шестнадцатиричной строки в бинарные данные
size_t HexStrToBinW(
  const wchar_t *s,
  void *pBuffer,
  size_t cbSize
  );
// Копирование строки в буфер
size_t StrCchCopyN(
  TCHAR *pDest,
  size_t cchDest,
  const TCHAR *pSrc,
  size_t cchSrc
  );
// Копирование строки в буфер
size_t StrCchCopy(
  TCHAR *pDest,
  size_t cchDest,
  const TCHAR *pSrc
  );
// Перемещение строки в буфер
size_t StrCchMoveN(
  TCHAR *pDest,
  size_t cchDest,
  const TCHAR *pSrc,
  size_t cchSrc
  );
// Перемещение строки в буфер
size_t StrCchMove(
  TCHAR *pDest,
  size_t cchDest,
  const TCHAR *pSrc
  );
// Заполнение буфера указанным символом
TCHAR *StrSet(
  TCHAR *s,
  TCHAR ch,
  size_t count
  );
// Удаление подстроки из строки
void StrDel(
  TCHAR *s,
  size_t index,
  size_t count
  );
// Создание дубликата строки
// (Возвращенный указатель необходимо освободить при помощи функции free)
TCHAR *StrDup(
  const TCHAR *s
  );
// Объединение строк
size_t StrCat(
  const TCHAR *s1,
  const TCHAR *s2,
  TCHAR *dest,
  size_t destlen
  );
//---------------------------------------------------------------------------
#ifdef __cplusplus
}  // extern "C"
#endif // __cplusplus
//---------------------------------------------------------------------------
#endif  // __STRUTILS_H__
