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
// ������� ��������� �������� � ������
char *SkipBlanksA(
  const char *s
  );
// ������� ��������� �������� � ������
wchar_t *SkipBlanksW(
  const wchar_t *s
  );
// ����� � ������ ������� �������
char *StrCharA(
  const char *s,
  char ch
  );
// ����� � ������ ������� �������
wchar_t *StrCharW(
  const wchar_t *s,
  wchar_t ch
  );
// ����� � ������ ������� ������� �� ������
char *StrCharSetA(
  const char *s,
  const char *cs
  );
// ����� � ������ ������� ������� �� ������
wchar_t *StrCharSetW(
  const wchar_t *s,
  const wchar_t *cs
  );
// ����� � ������ ���������� �������
char *StrRCharA(
  const char *s,
  char ch
  );
// ����� � ������ ���������� �������
wchar_t *StrRCharW(
  const wchar_t *s,
  wchar_t ch
  );
// ����� � ������ ���������� ������� �� ������
char *StrRCharSetA(
  const char *s,
  const char *cs
  );
// ����� � ������ ���������� ������� �� ������
wchar_t *StrRCharSetW(
  const wchar_t *s,
  const wchar_t *cs
  );
// ��������� ����� ��� ����� ��������
int StrCmpNIA(
  const char *s1,
  const char *s2,
  size_t count
  );
// ��������� ����� ��� ����� ��������
int StrCmpNIW(
  const wchar_t *s1,
  const wchar_t *s2,
  size_t count
  );
// ��������� ����� ��� ����� ��������
int StrCmpIA(
  const char *s1,
  const char *s2
  );
// ��������� ����� ��� ����� ��������
int StrCmpIW(
  const wchar_t *s1,
  const wchar_t *s2
  );
// �������������� ������ � ����������� �����
unsigned long StrToULNA(
  const char *s,
  size_t count,
  char **endptr,
  int base
  );
// �������������� ������ � ����������� �����
unsigned long StrToULNW(
  const wchar_t *s,
  size_t count,
  wchar_t **endptr,
  int base
  );
// �������������� ������ � ����������� �����
unsigned long StrToULA(
  const char *s,
  char **endptr,
  int base
  );
// �������������� ������ � ����������� �����
unsigned long StrToULW(
  const wchar_t *s,
  wchar_t **endptr,
  int base
  );
// �������������� ������ � ����������� �����
unsigned int StrToUIA(
  const char *s
  );
// �������������� ������ � ����������� �����
unsigned int StrToUIW(
  const wchar_t *s
  );
// �������������� ������������ ������ � ������
size_t UI64ToStrA(
  unsigned __int64 nValue,
  char chThousandSeparator,
  char *pBuffer
  );
// �������������� ������������ ������ � ������
size_t UI64ToStrW(
  unsigned __int64 nValue,
  wchar_t chThousandSeparator,
  wchar_t *pBuffer
  );
// �������������� �������� ������ � ����������������� ������
void BinToHexStrA(
  const void *pData,
  size_t cbData,
  char *pBuffer
  );
// �������������� �������� ������ � ����������������� ������
void BinToHexStrW(
  const void *pData,
  size_t cbData,
  wchar_t *pBuffer
  );
// �������������� ����������������� ����� � �����
int HexDigitValueA(
  char ch
  );
// �������������� ����������������� ����� � �����
int HexDigitValueW(
  wchar_t ch
  );
// �������������� ����������������� ������ � �������� ������
size_t HexStrToBinA(
  const char *s,
  void *pBuffer,
  size_t cbSize
  );
// �������������� ����������������� ������ � �������� ������
size_t HexStrToBinW(
  const wchar_t *s,
  void *pBuffer,
  size_t cbSize
  );
// ����������� ������ � �����
size_t StrCchCopyN(
  TCHAR *pDest,
  size_t cchDest,
  const TCHAR *pSrc,
  size_t cchSrc
  );
// ����������� ������ � �����
size_t StrCchCopy(
  TCHAR *pDest,
  size_t cchDest,
  const TCHAR *pSrc
  );
// ����������� ������ � �����
size_t StrCchMoveN(
  TCHAR *pDest,
  size_t cchDest,
  const TCHAR *pSrc,
  size_t cchSrc
  );
// ����������� ������ � �����
size_t StrCchMove(
  TCHAR *pDest,
  size_t cchDest,
  const TCHAR *pSrc
  );
// ���������� ������ ��������� ��������
TCHAR *StrSet(
  TCHAR *s,
  TCHAR ch,
  size_t count
  );
// �������� ��������� �� ������
void StrDel(
  TCHAR *s,
  size_t index,
  size_t count
  );
// �������� ��������� ������
// (������������ ��������� ���������� ���������� ��� ������ ������� free)
TCHAR *StrDup(
  const TCHAR *s
  );
// ����������� �����
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
